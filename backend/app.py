from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
from datetime import datetime
import os
import io
import pandas as pd
from excel_handler import ExcelHandler
from outlook_handler import crear_correo_outlook, WINDOWS_OUTLOOK_AVAILABLE
from closed_and_block_handler import ClosedAndBlockHandler

app = Flask(__name__, static_folder='static', template_folder='templates')
CORS(app)

# Configuration
DEFAULT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PlanillaEscalamientos.xlsx")
CLOSED_BLOCK_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ClosedAndBlock.xlsx")
excel = ExcelHandler(DEFAULT_PATH)
closed_block = ClosedAndBlockHandler(CLOSED_BLOCK_PATH)

# Almacenamiento en memoria para XOLUSAT (igual que st.session_state.xolusat)
xolusat_records = []


def es_sucursal(c):
    return "SUCURSAL" in c or c.startswith("SUC")


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


@app.route('/api/debug/contacts', methods=['GET'])
def debug_contacts():
    """Endpoint temporal para debuggear datos de custodios y contactos"""
    # Muestra custodios únicos del UNIFICADO
    custodios = {}
    for k, v in excel.data['unificado'].items():
        c = v.get('custodio', '')
        if c not in custodios:
            custodios[c] = 0
        custodios[c] += 1

    return jsonify({
        'contactos_semana_keys': list(excel.data['contactos'].keys()),
        'contactos_finde_keys': list(excel.data['contactos_finde'].keys()),
        'custodios_unicos': custodios,
        'unificado_ejemplo': {k: excel.data['unificado'][k] for k in list(excel.data['unificado'].keys())[:5]},
        'total_unificado': len(excel.data['unificado']),
        'total_contactos': len(excel.data['contactos']),
        'total_finde': len(excel.data['contactos_finde']),
        'total_suc': len(excel.data['contactos_suc'])
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
        # Limpiar líneas vacías
        lines = [line for line in content.split('\n') if line.strip()]
        clean_content = '\n'.join(lines)

        # Detectar si la primera fila es header (igual que original app2.py)
        primer_linea = clean_content.split("\n")[0].upper()
        has_header = "ID" in primer_linea or "ADDRESS" in primer_linea or "INICIO" in primer_linea

        if has_header:
            df = pd.read_csv(io.StringIO(clean_content), sep="\t")
        else:
            df = pd.read_csv(io.StringIO(clean_content), sep="\t", header=None)

        # Forzar nombres de columnas como strings "0", "1", "2"... para el frontend
        df.columns = [str(i) for i in range(len(df.columns))]

        # Reemplazar valores NaN por strings vacíos para evitar errores en JSON
        df = df.fillna("")

        # Agregar custodio desde UNIFICADO
        failures = []
        for _, row in df.iterrows():
            record = row.to_dict()
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

        # --- FLUJO ORIGINAL (app2.py líneas ~700-750) ---
        if id_norm in excel.data['unificado']:
            info = excel.data['unificado'][id_norm]
            custodio = info.get('custodio', '')
            cust_norm = excel.normalizar(custodio)

            # PRIORIZAR: Contactos específicos del ATM en contactos_suc SOLO para SUCURSAL
            # PERO si es fin de semana/feriado, buscar primero en contactos_finde para SUCURSAL
            if es_sucursal(cust_norm):
                if es_fin_de_semana and "SUCURSAL" in excel.data['contactos_finde']:
                    # Usar contacto de fin de semana para SUCURSAL
                    email, cc = excel.data['contactos_finde']["SUCURSAL"]
                elif id_norm in excel.data['contactos_suc']:
                    # Usar contacto específico del ATM
                    email, cc = excel.data['contactos_suc'][id_norm]
                else:
                    email = ""
                    cc = ""
            else:
                email = ""
                cc = ""
                
                # 1. Por ID exacto en contactos (semana o finde según corresponda)
                if id_norm in dict_contactos_activos:
                    email, cc = dict_contactos_activos[id_norm]
                    if not (isinstance(email, float) and pd.isna(email)):
                        pass
                    else:
                        email = ""

                # 2. Por Custodio exacto en contactos
                if not email and custodio in dict_contactos_activos:
                    email, cc = dict_contactos_activos[custodio]
                    if not (isinstance(email, float) and pd.isna(email)):
                        pass
                    else:
                        email = ""

                # 3. Búsqueda por custodio en contactos (match exacto primero, luego keyword)
                if not email:
                    cust_upper = cust_norm.upper().replace(" ", "")
                    for key, val in dict_contactos_activos.items():
                        key_upper = key.upper().replace(" ", "")
                        if cust_upper == key_upper:
                            email, cc = val
                            break
                    if not email:
                        for key, val in dict_contactos_activos.items():
                            key_upper = key.upper().replace(" ", "")
                            if key_upper.startswith("BHD"):
                                continue
                            if "BRINKS" in cust_upper and "BRINKS" in key_upper:
                                email, cc = val
                                break
                            elif "STE" in cust_upper and "STE" in key_upper:
                                email, cc = val
                                break
                            elif es_sucursal(cust_norm) and es_sucursal(key):
                                email, cc = val
                                break

                # 4. Si aún no tiene email y es SUCURSAL, buscar en contactos_suc
                if not email and es_sucursal(cust_norm) and id_norm in excel.data['contactos_suc']:
                    email, cc = excel.data['contactos_suc'][id_norm]

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
                <p>Atentamente,<br><b>Gestión de ATM</b></p>
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

    temp_path = os.path.join(os.path.dirname(DEFAULT_PATH), "temp_rcu.xlsx")
    file.save(temp_path)

    success, message = excel.procesar_rcu(temp_path)

    if os.path.exists(temp_path):
        os.remove(temp_path)

    if success:
        excel.cargar_datos()
        # message es un dict con actualizados, nuevos, etc.
        if isinstance(message, dict):
            return jsonify({'status': 'success', 'results': message})
        return jsonify({'status': 'success', 'message': str(message)})
    # message es un string de error
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
            return jsonify({'status': 'success', 'message': f'{incident} actualizado'})

    return jsonify({'status': 'error', 'message': 'Incidente no encontrado'}), 404


@app.route('/api/xolusat/clear', methods=['POST'])
def xolusat_clear():
    xolusat_records.clear()
    return jsonify({'status': 'success', 'message': 'Registros limpiados'})


# ==========================================
# CLOSED AND BLOCK
# ==========================================

@app.route('/api/closed-block/list', methods=['GET'])
def closed_block_list():
    return jsonify({'status': 'success', 'records': closed_block.listar_todos()})


@app.route('/api/closed-block/agregar', methods=['POST'])
def closed_block_agregar():
    data = request.json
    text = data.get('text', '')
    asunto = data.get('asunto', '')
    reportado_por = data.get('reportado_por', '')

    if not text:
        return jsonify({'status': 'error', 'message': 'No se proporcionaron IDs'}), 400

    # Parsear IDs del texto pegado
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    ids = []
    for line in lines:
        # Tomar solo la primera columna (tab-separated o coma-separated)
        parts = line.replace('\t', ',').split(',')
        id_val = parts[0].strip()
        if id_val and id_val.upper() not in ['ID', 'NAN', '']:
            ids.append(id_val)

    if not ids:
        return jsonify({'status': 'error', 'message': 'No se encontraron IDs válidos'}), 400

    result = closed_block.agregar_ids(ids, excel.data['unificado'], asunto, reportado_por)
    if 'error' in result:
        return jsonify({'status': 'error', 'message': result['error']}), 500

    return jsonify({
        'status': 'success',
        'added': result['added'],
        'skipped': result['skipped'],
        'total': result['total'],
        'message': f'{result["added"]} agregados, {result["skipped"]} ya existían. Total: {result["total"]}'
    })


@app.route('/api/closed-block/buscar', methods=['POST'])
def closed_block_buscar():
    data = request.json
    
    # Acepta texto pegado (con formato) o array de IDs
    text = data.get('text', '')
    ids_raw = data.get('ids', [])
    
    if text:
        # Parsear IDs del texto pegado (igual que agregar)
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        ids_parsed = []
        for line in lines:
            parts = line.replace('\t', ',').split(',')
            id_val = parts[0].strip()
            if id_val and id_val.upper() not in ['ID', 'NAN', '']:
                ids_parsed.append(id_val)
        results = closed_block.buscar_ids(ids_parsed)
    elif ids_raw:
        results = closed_block.buscar_ids(ids_raw)
    else:
        results = []
    
    return jsonify({'status': 'success', 'found': results})


@app.route('/api/closed-block/eliminar', methods=['POST'])
def closed_block_eliminar():
    data = request.json
    atm_id = data.get('id', '')
    if not atm_id:
        return jsonify({'status': 'error', 'message': 'ID requerido'}), 400
    success = closed_block.eliminar_id(atm_id)
    if success:
        return jsonify({'status': 'success', 'message': 'Eliminado'})
    return jsonify({'status': 'error', 'message': 'No se encontró'}), 400


@app.route('/api/closed-block/limpiar', methods=['POST'])
def closed_block_limpiar():
    removed, remaining = closed_block.limpiar_vencidos()
    return jsonify({
        'status': 'success',
        'message': f'{removed} registros vencidos eliminados. Restan: {remaining}'
    })


@app.route('/api/generate-scripts-cb', methods=['POST'])
def generate_scripts_cb():
    data = request.json
    text = data.get('text', '')
    
    if not text:
        return jsonify({'status': 'error', 'message': 'Sin datos'}), 400
    
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    ids_data = {}
    for line in lines:
        parts = [p.strip() for p in line.replace('\t', ',').split(',')]
        
        id_val = parts[0].strip().upper() if parts[0] else ''
        if not id_val or id_val.upper() in ['ID', 'NAN', '']:
            continue
            
        id_norm = closed_block.normalizar(id_val)
        
        tk_val = parts[10].strip() if len(parts) > 10 else ''
        if not tk_val and len(parts) > 9:
            tk_val = parts[9].strip()
        
        if tk_val and not tk_val.isdigit():
            tk_val = ''
        
        if id_norm not in ids_data:
            ids_data[id_norm] = []
        if tk_val:
            ids_data[id_norm].append(tk_val)
    
    if not ids_data:
        return jsonify({'status': 'error', 'message': 'No se encontraron IDs'}), 400
    
    ids_raw = list(ids_data.keys())
    found = closed_block.buscar_ids(ids_raw)
    
    scripts = []
    for r in found:
        id_norm = r['id']
        tks = [t for t in ids_data.get(id_norm, []) if t and t != 'N/A']
        asunto = r.get('asunto', 'N/A')
        reportado_por = r.get('reportado_por', 'SUCURSAL')
        
        for tk in tks:
            script_line = f"#07# ATM bloqueado por cliente // {asunto} // {reportado_por}"
            scripts.append({"ticket": tk, "comentario": script_line})
    
    return jsonify({'status': 'success', 'scripts': scripts})


@app.route('/api/shutdown', methods=['POST'])
def shutdown_server():
    import os
    print("\n" + "=" * 40)
    print("  USUARIO CERRO LA APLICACION")
    print("  Servidor detenido.")
    print("=" * 40)
    
    if os.path.exists('server.pid'):
        os.remove('server.pid')
    
    return jsonify({'status': 'success'})


# ==========================================
# MAIN
# ==========================================

running = True

def signal_handler(sig, frame):
    global running
    running = False
    print("\nServidor detenido.")

if __name__ == '__main__':
    import signal
    signal.signal(signal.SIGINT, signal_handler)
    
    print("=" * 50)
    print("  ESCALAMIENTOS APP - Servidor iniciado")
    print("  http://localhost:5000")
    print("=" * 50)
    excel.cargar_datos()
    
    import os
    with open('server.pid', 'w') as f:
        f.write(str(os.getpid()))
    
    app.run(debug=False, port=5000)
