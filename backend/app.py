from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from datetime import datetime
import os
import sys
import io
import json
import pandas as pd
from excel_handler import ExcelHandler
from outlook_handler import crear_correo_outlook, WINDOWS_OUTLOOK_AVAILABLE

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)

# Configuration
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
    # Primer ejecucion: copiar xlsx desde el bundle
    xlsx_destino = os.path.join(APP_DIR, "PlanillaEscalamientos.xlsx")
    xlsx_bundle = os.path.join(sys._MEIPASS, "PlanillaEscalamientos.xlsx")
    if not os.path.exists(xlsx_destino) and os.path.exists(xlsx_bundle):
        import shutil
        shutil.copy2(xlsx_bundle, xlsx_destino)
    # Crear acceso directo
    try:
        import tempfile, subprocess
        ps_script = (
            '$desktop = [Environment]::GetFolderPath("Desktop"); '
            '$sc_path = Join-Path $desktop "EscalamientosApp.lnk"; '
            '$ws = New-Object -ComObject WScript.Shell; '
            '$sc = $ws.CreateShortcut($sc_path); '
            '$sc.TargetPath = "' + sys.executable + '"; '
            '$sc.WorkingDirectory = "' + APP_DIR + '"; '
            '$sc.Description = "Gestion de Escalamientos ATM - BHD"; '
            '$sc.Save()'
        )
        ps_file = os.path.join(tempfile.gettempdir(), "_esc_shortcut.ps1")
        with open(ps_file, 'w', encoding='utf-8') as f:
            f.write(ps_script)
        subprocess.run(
            ['powershell', '-ExecutionPolicy', 'Bypass', '-File', ps_file],
            capture_output=True,
            creationflags=0x08000000  # CREATE_NO_WINDOW
        )
        try: os.remove(ps_file)
        except: pass
    except:
        pass
    DEFAULT_PATH = xlsx_destino
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))
    DEFAULT_PATH = os.path.join(APP_DIR, "PlanillaEscalamientos.xlsx")
excel = ExcelHandler(DEFAULT_PATH)

# ── Persistencia XOLUSAT ──────────────────────────────────────────────────────
XOLUSAT_FILE = os.path.join(APP_DIR, 'xolusat_records.json')

def _cargar_xolusat():
    """Carga registros XOLUSAT desde archivo JSON. Devuelve lista vacía si no existe o hay error."""
    if os.path.exists(XOLUSAT_FILE):
        try:
            with open(XOLUSAT_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Warning: no se pudo leer xolusat_records.json: {e}")
    return []

def _guardar_xolusat():
    """Persiste xolusat_records en archivo JSON."""
    try:
        with open(XOLUSAT_FILE, 'w', encoding='utf-8') as f:
            json.dump(xolusat_records, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Warning: no se pudo guardar xolusat_records.json: {e}")

xolusat_records = _cargar_xolusat()


def es_sucursal(c):
    return "SUCURSAL" in c or c.startswith("SUC")


def _limpiar_par_email(email, cc):
    """Devuelve ('', '') si el email es NaN de pandas, o el par original."""
    if isinstance(email, float) and pd.isna(email):
        return "", ""
    return email, cc


def obtener_contacto_atm(id_norm, custodio, es_fin_de_semana, data):
    """
    Resuelve (email, cc) para un ATM dado su custodio y el contexto temporal.

    Orden de prioridad:
      SUCURSAL  → finde/feriado: contactos_finde["SUCURSAL"]
                → semana:        contactos_suc[id_norm]
      TERCEROS  → 1. ID exacto en dict activo
                → 2. Nombre de custodio exacto en dict activo
                → 3. Match exacto (sin espacios) en dict activo
                → 4. Keyword (BRINKS / STE) en dict activo
    """
    cust_norm = excel.normalizar(custodio)
    dict_contactos = data['contactos_finde'] if es_fin_de_semana else data['contactos']

    # ── SUCURSAL ─────────────────────────────────────────────────────────────
    if es_sucursal(cust_norm):
        if es_fin_de_semana and "SUCURSAL" in data['contactos_finde']:
            return _limpiar_par_email(*data['contactos_finde']["SUCURSAL"])
        if id_norm in data['contactos_suc']:
            return _limpiar_par_email(*data['contactos_suc'][id_norm])
        return "", ""

    # ── TERCEROS ─────────────────────────────────────────────────────────────
    # 1. Por ID exacto
    if id_norm in dict_contactos:
        email, cc = _limpiar_par_email(*dict_contactos[id_norm])
        if email:
            return email, cc

    # 2. Por nombre de custodio exacto
    if custodio in dict_contactos:
        email, cc = _limpiar_par_email(*dict_contactos[custodio])
        if email:
            return email, cc

    # 3 y 4. Búsqueda heurística (match sin espacios → keywords)
    cust_upper = cust_norm.upper().replace(" ", "")
    match_exacto = None
    match_keyword = None

    for key, val in dict_contactos.items():
        key_upper = key.upper().replace(" ", "")

        # Match exacto: siempre se chequea, incluso para claves con prefijo BHD
        # (ej: "BHD - STE Metro" normalizado = "BHDSTEMTERO" debe matchear exacto)
        if match_exacto is None and cust_upper == key_upper:
            match_exacto = val

        # Keyword heurístico: se omiten claves BHD para evitar falsos positivos
        # (ej: "BRINKSESTE" contiene "STE" pero no es contacto de STE)
        if key_upper.startswith("BHD"):
            continue
        if match_keyword is None:
            if "BRINKS" in cust_upper and "BRINKS" in key_upper:
                match_keyword = val
            elif "STE" in cust_upper and "STE" in key_upper:
                match_keyword = val

    for candidato in (match_exacto, match_keyword):
        if candidato is not None:
            email, cc = _limpiar_par_email(*candidato)
            if email:
                return email, cc

    return "", ""


# ==========================================
# RUTAS PRINCIPALES
# ==========================================

@app.route('/')
def index():
    es_domingo = datetime.now().weekday() == 6
    return render_template('index.html', es_domingo=es_domingo)


@app.route('/api/status', methods=['GET'])
def get_status():
    excel_exists = os.path.exists(DEFAULT_PATH)
    n_atms = len(excel.data['unificado']) if excel_exists else 0
    es_domingo = datetime.now().weekday() == 6
    return jsonify({
        'excel_connected': excel_exists,
        'n_atms': n_atms,
        'outlook_available': WINDOWS_OUTLOOK_AVAILABLE,
        'es_domingo': es_domingo
    })


@app.route('/api/load-data', methods=['GET'])
def load_data():
    success, message = excel.cargar_datos()
    if success:
        return jsonify({
            'status': 'success',
            'message': message,
            'data': excel.data
        })
    return jsonify({'status': 'error', 'message': message}), 400


@app.route('/api/process-failures', methods=['POST'])
def process_failures():
    content = request.json.get('text', '')
    if not content or not content.strip():
        return jsonify({'status': 'error', 'message': 'No se proporcionaron datos.'}), 400

    try:
        # Limpiar líneas vacías al inicio y final
        lines = [line for line in content.split('\n') if line.strip()]
        clean_content = '\n'.join(lines)

        # Siempre procesamos como SIN header para no perder la primera fila
        # El usuario siempre pega CON encabezados según la nueva instrucción
        # Por eso usamos header=None para que todo sea data
        df = pd.read_csv(io.StringIO(clean_content), sep="\t", header=None)

        # Forzar nombres de columnas como strings "0", "1", "2"... para el frontend
        df.columns = [str(i) for i in range(len(df.columns))]

        # Reemplazar valores NaN por strings vacíos para evitar errores en JSON
        df = df.fillna("")

        # Detectar si la primera fila es header para excluirla del procesamiento
        def es_fila_header(row):
            """Detecta si una fila parece ser un header (encabezado de Excel)"""
            primer_valor = str(row.get('0', '')).upper().strip()
            header_keywords = ['ID', 'ADDRESS', 'INICIO', 'MODEL', 'FECHA', 'DESCRIPTION', 
                               'TICKET', 'AGENCIA', 'TIPO', 'STATUS', 'STATE', 'NOMBRE']
            return any(kw in primer_valor for kw in header_keywords)

        # Agregar custodio desde UNIFICADO
        failures = []
        for idx, row in enumerate(df.iterrows()):
            record = row[1].to_dict()
            
            # Marcar la primera fila como header si corresponde
            if idx == 0 and es_fila_header(record):
                record['_is_header'] = True
            else:
                record['_is_header'] = False
            
            id_raw = str(record.get('0', '')).strip()
            id_norm = excel.normalizar(id_raw)
            info = excel.data['unificado'].get(id_norm, {})
            record['_custodio'] = info.get('custodio', '')
            record['_found'] = bool(info)
            failures.append(record)

        return jsonify({
            'status': 'success',
            'failures': failures
        })
    except Exception as e:
        return jsonify({'status': 'error', 'message': f"Error de formato: {str(e)}"}), 400


# ==========================================
# ENVIAR CORREOS - Lógica idéntica al original app2.py
# ==========================================

@app.route('/api/send-emails', methods=['POST'])
def send_emails():
    data = request.json
    failures = data.get('failures', [])
    is_feriado = data.get('is_feriado', False)

    if not failures:
        return jsonify({'status': 'error', 'message': 'No hay fallas para procesar.'}), 400

    # Detectar si es domingo (igual que original)
    es_domingo = datetime.now().weekday() == 6
    es_fin_de_semana = es_domingo or is_feriado

    # Elegir diccionario de contactos según el día (original: semana vs finde)
    if es_fin_de_semana:
        dict_contactos_activos = excel.data['contactos_finde']
    else:
        dict_contactos_activos = excel.data['contactos']

    df = pd.DataFrame(failures)
    df.columns = [str(i) for i in range(len(df.columns))]
    df['id_norm'] = df['0'].apply(excel.normalizar)
    grupos = df.groupby('id_norm')

    results = {'abiertos': 0, 'sin_sla': 0, 'sin_contacto': 0}

    for id_norm, group in grupos:
        if not id_norm.strip() or id_norm.lower() == "nan":
            continue

        primer_fila = group.iloc[0]
        id_raw = str(primer_fila['0'])
        email = ""
        cc = ""
        asunto = ""

        # ── Resolver contacto ─────────────────────────────────────────────────
        if id_norm in excel.data['unificado']:
            info = excel.data['unificado'][id_norm]
            custodio = info.get('custodio', '')
            email, cc = obtener_contacto_atm(id_norm, custodio, es_fin_de_semana, excel.data)

            if not email or (isinstance(email, float) and pd.isna(email)):
                results['sin_contacto'] += 1
                continue

            tipo_pasted = str(primer_fila['2']) if '2' in primer_fila else ""
            asunto = f"ESCALAMIENTO FALLA- ATM {id_raw} - {tipo_pasted}"

        else:
            # ATM NO está en unificado → buscar en dict_suc directamente (igual que original)
            if id_norm in excel.data['contactos_suc']:
                email, cc = excel.data['contactos_suc'][id_norm]
                tipo_pasted = str(primer_fila['2']) if '2' in primer_fila else ""
                asunto = f"ESCALAMIENTO FALLA- ATM {id_raw} - {tipo_pasted}"
            else:
                results['sin_sla'] += 1
                continue

        if email and not (isinstance(email, float) and pd.isna(email)):
            # --- GENERACIÓN DE CUERPO HTML ---
            filas_html = ""
            for _, row in group.iterrows():
                try:
                    f_id = str(row['0'])
                    f_agencia = str(row['2']) if '2' in row else ""
                    f_modelo = str(row['4']) if '4' in row else ""
                    f_fecha = str(row['5']) if '5' in row else ""
                    f_desc = str(row['6']) if '6' in row else ""
                    f_ticket = str(row['9']) if '9' in row else "N/A"

                    filas_html += f"""
                    <tr style='border-bottom: 1px solid #eee;'>
                        <td style='padding: 10px; border: 1px solid #ddd;'>{f_id}</td>
                        <td style='padding: 10px; border: 1px solid #ddd;'>{f_agencia}</td>
                        <td style='padding: 10px; border: 1px solid #ddd;'>{f_modelo}</td>
                        <td style='padding: 10px; border: 1px solid #ddd;'>{f_fecha}</td>
                        <td style='padding: 10px; border: 1px solid #ddd;'>{f_desc}</td>
                        <td style='padding: 10px; border: 1px solid #ddd;'><b>{f_ticket}</b></td>
                    </tr>
                    """
                except (IndexError, KeyError):
                    continue

            cuerpo = f"""
            <div style='font-family: Calibri, Arial, sans-serif; font-size: 15px; color: #333;'>
                <p>Estimados,</p>
                <p>Se observan las siguientes fallas en el ATM <b>{id_raw}</b>:</p>
                <p><b>Dirección:</b> {group.iloc[0, 1] if len(group.columns) > 1 else 'N/D'}</p>
                <table style='border-collapse: collapse; width: 100%; border: 1px solid #ccc; font-size: 13px;'>
                    <thead style='background-color: #54b948; color: white;'>
                        <tr>
                            <th style='padding: 10px; border: 1px solid #54b948;'>ID</th>
                            <th style='padding: 10px; border: 1px solid #54b948;'>Agencia</th>
                            <th style='padding: 10px; border: 1px solid #54b948;'>Modelo</th>
                            <th style='padding: 10px; border: 1px solid #54b948;'>Fecha Falla</th>
                            <th style='padding: 10px; border: 1px solid #54b948;'>Descripción Falla</th>
                            <th style='padding: 10px; border: 1px solid #54b948;'>Ticket</th>
                        </tr>
                    </thead>
                    <tbody>{filas_html}</tbody>
                </table>
                <br>
                <p>Atentamente</p>
            </div>
            """

            success, msg = crear_correo_outlook(email, cc, asunto, cuerpo)
            if success:
                results['abiertos'] += 1
            else:
                return jsonify({'status': 'error', 'message': msg}), 500

    return jsonify({'status': 'success', 'results': results})


# ==========================================
# GENERAR SCRIPTS - Backend (igual que original app2.py tab2)
# ==========================================

@app.route('/api/generate-scripts', methods=['POST'])
def generate_scripts():
    data = request.json
    failures = data.get('failures', [])
    is_feriado = data.get('is_feriado', False)

    if not failures:
        return jsonify({'status': 'error', 'message': 'No hay fallas para procesar.'}), 400

    es_domingo = datetime.now().weekday() == 6
    es_fin_de_semana = es_domingo or is_feriado

    scripts = []

    for row in failures:
        id_raw = str(row.get('0', ''))
        if not id_raw.strip() or id_raw.lower() == "nan":
            continue

        id_norm = excel.normalizar(id_raw)
        info = excel.data['unificado'].get(id_norm, {})
        custodio = info.get('custodio', 'N/A')
        cust_norm = excel.normalizar(custodio)

        tipo_pasted = str(row.get('2', ''))
        asunto = f"ESCALAMIENTO FALLA- ATM {id_raw} - {tipo_pasted}"

        try:
            falla_desc = str(row.get('6', ''))
            ticket = str(row.get('9', 'N/A'))

            # Determinar destino (igual que original app2.py tab2)
            if "BRINK" in cust_norm:
                destino_script = "Brinks"
            elif "STE" in cust_norm or "SERVICIO TECNICO" in cust_norm:
                destino_script = "STE"
            else:
                destino_script = custodio

            # Si finde/feriado, sobreescribir destino con dict_finde (igual que original)
            if es_fin_de_semana:
                found_finde = False
                for key_finde in excel.data['contactos_finde']:
                    key_finde_norm = excel.normalizar(key_finde)
                    if key_finde_norm != "SUCURSAL" and (key_finde_norm in cust_norm or cust_norm in key_finde_norm):
                        if "BRINK" in key_finde_norm:
                            destino_script = "Brinks"
                        elif "STE" in key_finde_norm or "SERVICIO TECNICO" in key_finde_norm:
                            destino_script = "STE"
                        else:
                            destino_script = key_finde
                        found_finde = True
                        break
                if not found_finde and "SUCURSAL" in excel.data['contactos_finde']:
                    destino_script = "SUCURSAL"

            script_line = f"#15# Se escala a {destino_script} {asunto} + {falla_desc}"
            scripts.append({"ticket": ticket, "comentario": script_line})

        except Exception:
            continue

    return jsonify({'status': 'success', 'scripts': scripts})


# ==========================================
# EXPORTAR SCRIPTS A EXCEL
# ==========================================

@app.route('/api/export-scripts', methods=['POST'])
def export_scripts():
    data = request.json
    scripts = data.get('scripts', [])
    if not scripts:
        return jsonify({'status': 'error', 'message': 'No hay datos para exportar'}), 400

    rows = []
    for s in scripts:
        tk = s.get('ticket', '')
        comentario = s.get('comentario', '')
        if comentario.startswith(tk + ' '):
            comentario = comentario[len(tk) + 1:]
        rows.append({'TK': tk, 'COMENTARIOS': comentario})
    
    df = pd.DataFrame(rows)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Tickets')

    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f"Scripts_Escalamiento_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    )


# ==========================================
# ATMs NO ENCONTRADOS + ALTA
# ==========================================

@app.route('/api/check-atm', methods=['POST'])
def check_atm():
    data = request.json
    atm_id = data.get('id', '')
    if not atm_id:
        return jsonify({'status': 'error', 'message': 'ID requerido'}), 400

    id_norm = excel.normalizar(atm_id)
    if id_norm in excel.data['unificado']:
        return jsonify({'status': 'found', 'info': excel.data['unificado'][id_norm]})
    return jsonify({'status': 'not_found'})


@app.route('/api/add-atm', methods=['POST'])
def add_atm():
    data = request.json
    atm_id = data.get('id')
    nombre = data.get('nombre')
    sla = data.get('sla')
    custodio = data.get('custodio')

    if not all([atm_id, nombre, sla, custodio]):
        return jsonify({'status': 'error', 'message': 'Todos los campos son obligatorios.'}), 400

    # Guardar en Excel
    success, message = excel.guardar_atm(atm_id, nombre, sla, custodio)
    if success:
        # Actualizar datos en memoria (igual que original)
        id_norm = excel.normalizar(atm_id)
        excel.data['unificado'][id_norm] = {
            'nombre': nombre,
            'custodio': custodio,
            'sla_marcas': sla,
            'sla_brinks': '',
            'denominacion': '',
            'zona': '',
            'disp_o_mult': '',
            'address2': nombre,
            'city': '',
            'ip_address': '',
            'district': sla
        }
        return jsonify({'status': 'success', 'message': message})
    return jsonify({'status': 'error', 'message': message}), 400


# ==========================================
# SUBIR Y PROCESAR RCU
# ==========================================

@app.route('/api/upload-rcu', methods=['POST'])
def upload_rcu():
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'No selected file'}), 400

    # Leer directo del stream sin guardar a disco
    import io
    data = file.read()
    df_nuevo = pd.read_excel(io.BytesIO(data), header=2)
    
    success, message = excel.procesar_rcu_desde_df(df_nuevo)
    
    if success:
        excel.cargar_datos()
        if isinstance(message, dict):
            return jsonify({'status': 'success', 'results': message})
        return jsonify({'status': 'success', 'message': str(message)})
    return jsonify({'status': 'error', 'message': str(message)}), 400


# ==========================================
# XOLUSAT
# ==========================================

@app.route('/api/xolusat/search', methods=['POST'])
def xolusat_search():
    data = request.json
    atm_id = data.get('id', '')
    id_norm = excel.normalizar(atm_id)

    if id_norm in excel.data['unificado']:
        info = excel.data['unificado'][id_norm]
        return jsonify({
            'status': 'found',
            'nombre': info.get('nombre', ''),
            'sla': info.get('sla_marcas', ''),
            'custodio': info.get('custodio', '')
        })
    return jsonify({'status': 'not_found'})


@app.route('/api/xolusat/register', methods=['POST'])
def xolusat_register():
    data = request.json
    incident = data.get('incident', '')
    estado = data.get('estado', '')
    id_atm = data.get('id_atm', '')
    subcategoria = data.get('subcategoria', '')
    detalle = data.get('detalle', '')
    sla = data.get('sla', '')
    atm_nombre = data.get('atm_nombre', '')
    custodio = data.get('custodio', '')
    send_email = data.get('send_email', False)

    if not incident or not id_atm:
        return jsonify({'status': 'error', 'message': 'Incident e ID ATM son obligatorios.'}), 400

    registro = {
        'incident': incident,
        'estado': estado,
        'id_atm': id_atm.upper(),
        'subcategoria': subcategoria,
        'detalle': detalle,
        'sla': sla,
        'atm_nombre': atm_nombre,
        'custodio': custodio,
        'fecha_reg': datetime.now().strftime("%d/%m/%Y %H:%M")
    }
    xolusat_records.append(registro)
    _guardar_xolusat()

    if send_email:
        to_xol = "operaciones@xolusat.com; imartinez@xolusat.com; fmella@xolusat.com; centrodecontrol@xolusat.com"
        cc_xol = "IM.BHD@ncratleos.com; gestion_atm@bhd.com.do; Aurora.Rodriguez@ncratleos.com; Daisy.Perez@ncratleos.com; Alexander.Torres@ncratleos.com; Monica.Serrano@ncratleos.com"
        asunto_xol = f"ESCALAMIENTO - ATM {id_atm.upper()} {atm_nombre}"

        cuerpo_xol = f"""
        <div style='font-family: Arial, sans-serif; font-size: 14px;'>
            <p>Estimados,</p>
            <p>Su apoyo con la atención, ATM se encuentra <b>{subcategoria}</b>:</p>
            <p><span style='background-color: yellow;'><b>Favor compartir ETA y actualizar en D1.</b></span></p>
            <p><b>Notificación de Creación de Incidente</b></p>
            <br>
            <b>Incident:</b> {incident}<br>
            <b>Estado:</b> {estado}<br>
            <b>Categoría:</b> {id_atm.upper()}<br>
            <b>Subcategoría:</b> {subcategoria}<br>
            <b>Detalle:</b> {detalle}<br>
            <b>SLA:</b> {sla}<br>
            <b>ATM:</b> {atm_nombre}<br>
            <br>
            <p>Quedamos atentos a cualquier novedad.</p>
            <p>Saludos.</p>
        </div>
        """

        success, msg = crear_correo_outlook(to_xol, cc_xol, asunto_xol, cuerpo_xol)
        if success:
            return jsonify({'status': 'success', 'message': f'Correo {incident} abierto', 'registro': registro})
        return jsonify({'status': 'warning', 'message': 'Registro guardado pero no se pudo abrir Outlook', 'registro': registro})

    return jsonify({'status': 'success', 'message': f'Registro {incident} guardado', 'registro': registro})


@app.route('/api/xolusat/list', methods=['GET'])
def xolusat_list():
    estado_filter = request.args.get('estado', '')
    subcat_filter = request.args.get('subcategoria', '')

    filtered = xolusat_records
    if estado_filter and estado_filter != 'Todos':
        filtered = [r for r in filtered if r['estado'] == estado_filter]
    if subcat_filter and subcat_filter != 'Todas':
        filtered = [r for r in filtered if r['subcategoria'] == subcat_filter]

    return jsonify({'status': 'success', 'records': filtered})


@app.route('/api/xolusat/update-status', methods=['POST'])
def xolusat_update_status():
    data = request.json
    incident = data.get('incident', '')
    nuevo_estado = data.get('estado', '')

    for item in xolusat_records:
        if item['incident'] == incident:
            item['estado'] = nuevo_estado
            _guardar_xolusat()
            return jsonify({'status': 'success', 'message': f'{incident} actualizado'})

    return jsonify({'status': 'error', 'message': 'Incidente no encontrado'}), 404


@app.route('/api/xolusat/clear', methods=['POST'])
def xolusat_clear():
    xolusat_records.clear()
    _guardar_xolusat()
    return jsonify({'status': 'success', 'message': 'Registros limpiados'})


@app.route('/api/shutdown', methods=['POST'])
def shutdown_server():
    import os, signal
    print("\n" + "=" * 40)
    print("  USUARIO CERRO LA APLICACION")
    print("  Servidor detenido.")
    print("=" * 40)

    if os.path.exists('server.pid'):
        os.remove('server.pid')

    os.kill(os.getpid(), signal.SIGTERM)
    return jsonify({'status': 'success'})


# ==========================================
# TAB CONTACTOS
# ==========================================

@app.route('/api/contactos/list', methods=['GET'])
def contactos_list():
    result = excel.obtener_contactos_custodio()
    if 'error' in result:
        return jsonify({'status': 'error', 'message': result['error']}), 500
    return jsonify({'status': 'success', 'data': result})

@app.route('/api/contactos/guardar', methods=['POST'])
def contactos_guardar():
    data = request.json
    custodio = data.get('custodio', '').strip()
    email = data.get('email', '').strip()
    cc = data.get('cc', '').strip()
    aplica_finde = data.get('aplica_finde', False)
    email_finde = data.get('email_finde', '').strip()
    cc_finde = data.get('cc_finde', '').strip()
    solo = data.get('solo', '').strip()
    tipo = data.get('tipo', 'tercero').strip()

    if not custodio:
        return jsonify({'status': 'error', 'message': 'Falta custodio'}), 400

    result = excel.actualizar_contactos_custodio(custodio, email, cc, aplica_finde, tipo, email_finde, cc_finde, solo)
    if 'error' in result:
        return jsonify({'status': 'error', 'message': result['error']}), 500

    excel.cargar_datos()
    return jsonify(result)



# ==========================================
# MAIN
# ==========================================

if __name__ == '__main__':
    print("=" * 50)
    print("  ESCALAMIENTOS APP - Servidor iniciado")
    print("  http://localhost:5000")
    print("=" * 50)
    excel.cargar_datos()

    app.run(debug=False, host='127.0.0.1', port=5000)
