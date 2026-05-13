import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime

class ExcelHandler:
    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.data = {
            'contactos': {},
            'contactos_suc': {},
            'contactos_finde': {},
            'unificado': {}
        }

    def normalizar(self, s):
        if pd.isna(s):
            return ""
        t = str(s).upper().strip()
        for char in [".", "-", "_", " ", "/"]:
            t = t.replace(char, "")
        return t

    def normalizar_custodio(self, region, zone):
        zone = str(zone).strip().upper() if pd.notna(zone) else ''
        if 'Off-prime Brinks' in region: return f'Brinks {zone}'
        if 'Off-prime STE' in region: return 'BHD - STE Metro'
        if region == 'Sucursal-Sucursal': return 'SUCURSAL'
        if 'Sucursal-DriveUP Brinks' in region: return f'Brinks {zone}'
        if 'Sucursal-DriveUP STE' in region: return 'BHD - STE Metro'
        if 'Sucursal-Sucursal STE' in region: return 'BHD - STE Metro'
        if region in ['SUCURSAL', 'BHD - STE Metro', 'Brinks METRO', 'Brinks NORTE', 'Brinks ESTE', 'Brinks SUR']:
            return region
        return region

    def cargar_datos(self):
        if not os.path.exists(self.excel_path):
            return False, "Planilla no encontrada"
        
        try:
            xls = pd.ExcelFile(self.excel_path)
            
            # 1. CONTACTOS SEMANA
            if "CONTACTOS SEMANA" in xls.sheet_names:
                df_semana = pd.read_excel(xls, "CONTACTOS SEMANA")
                self.data['contactos'] = {
                    self.normalizar(r.iloc[0]): [r.iloc[1], r.iloc[2]]
                    for _, r in df_semana.iterrows() if not pd.isna(r.iloc[0])
                }

            # 2. CONTACTOS_SUC
            if "CONTACTOS_SUC" in xls.sheet_names:
                df_suc = pd.read_excel(xls, "CONTACTOS_SUC")
                for _, row in df_suc.iterrows():
                    if len(row) > 0 and pd.notna(row.iloc[0]):
                        sid = self.normalizar(row.iloc[0])
                        semail = row.iloc[6] if len(row) > 6 else ""
                        scopia = row.iloc[7] if len(row) > 7 else ""
                        self.data['contactos_suc'][sid] = [semail, scopia]

            # 3. CONTACTOS FINDE
            if "CONTACTOS FINDE" in xls.sheet_names:
                df_finde = pd.read_excel(xls, "CONTACTOS FINDE")
                for _, row in df_finde.iterrows():
                    reg = str(row.get('REGIONES', "")).strip()
                    if reg:
                        self.data['contactos_finde'][self.normalizar(reg)] = [row.get('CONTACTOS', ""), row.get('COPIA', "")]

            # 4. UNIFICADO
            if "UNIFICADO" in xls.sheet_names:
                df_unif = pd.read_excel(xls, "UNIFICADO")
                for _, row in df_unif.iterrows():
                    if pd.notna(row.iloc[0]):
                        id_n = self.normalizar(str(row.iloc[0]))
                        self.data['unificado'][id_n] = {
                            'nombre': str(row.iloc[1]) if len(row) > 1 else "",
                            'custodio': str(row.iloc[2]) if len(row) > 2 else "",
                            'sla_marcas': str(row.iloc[3]) if len(row) > 3 else "",
                            'sla_brinks': str(row.iloc[4]) if len(row) > 4 else "",
                            'denominacion': str(row.iloc[5]) if len(row) > 5 else "",
                            'zona': str(row.iloc[6]) if len(row) > 6 else "",
                            'disp_o_mult': str(row.iloc[7]) if len(row) > 7 else "",
                            'address2': str(row.iloc[8]) if len(row) > 8 else "",
                            'city': str(row.iloc[9]) if len(row) > 9 else "",
                            'ip_address': str(row.iloc[10]) if len(row) > 10 else "",
                            'district': str(row.iloc[11]) if len(row) > 11 else ""
                        }
            return True, "Datos cargados correctamente"
        except Exception as e:
            return False, str(e)

    def guardar_atm(self, atm_id, nombre, sla, custodio):
        try:
            wb = load_workbook(self.excel_path)
            # RCU sheet
            if "RCU" in wb.sheetnames:
                ws_rcu = wb["RCU"]
                new_row = ws_rcu.max_row + 1
                ws_rcu.cell(row=new_row, column=1, value=atm_id.upper())
                ws_rcu.cell(row=new_row, column=3, value=nombre)
                ws_rcu.cell(row=new_row, column=12, value=sla)
            # SLA sheet
            if "SLA" in wb.sheetnames:
                ws_sla = wb["SLA"]
                new_row = ws_sla.max_row + 1
                ws_sla.cell(row=new_row, column=1, value=atm_id.upper())
                ws_sla.cell(row=new_row, column=4, value=custodio)
                ws_sla.cell(row=new_row, column=8, value=sla)
            
            # Update UNIFICADO sheet as well
            if "UNIFICADO" in wb.sheetnames:
                ws_unif = wb["UNIFICADO"]
                new_row = ws_unif.max_row + 1
                ws_unif.cell(row=new_row, column=1, value=atm_id.upper())
                ws_unif.cell(row=new_row, column=2, value=nombre)
                ws_unif.cell(row=new_row, column=3, value=custodio)
                ws_unif.cell(row=new_row, column=4, value=sla)
                ws_unif.cell(row=new_row, column=12, value=sla)

            wb.save(self.excel_path)
            return True, "ATM guardado"
        except Exception as e:
            return False, str(e)

    def procesar_rcu(self, ruta_rcu_nuevo):
        try:
            df_nuevo = pd.read_excel(ruta_rcu_nuevo, header=2)
            wb = load_workbook(self.excel_path)
            
            # Ensure UNIFICADO exists
            if "UNIFICADO" not in wb.sheetnames:
                ws_unif = wb.create_sheet("UNIFICADO")
                headers = ["ID", "NOMBRE", "CUSTODIO", "SLA_MARCAS", "SLA_BRINKS", "DENOMINACION", "ZONA", "DISP_O_MULT", "ADDRESS2", "CITY", "IP_ADDRESS", "DISTRICT"]
                for i, h in enumerate(headers, 1): ws_unif.cell(row=1, column=i, value=h)
            else:
                ws_unif = wb["UNIFICADO"]

            # Map existing IDs
            ids_unif = {}
            for row in range(2, ws_unif.max_row + 1):
                cell_val = ws_unif.cell(row=row, column=1).value
                if cell_val: ids_unif[str(cell_val).strip().upper()] = row

            ws_rcu = wb["RCU"] if "RCU" in wb.sheetnames else None
            if not ws_rcu: return False, "No se encontró hoja RCU"

            ids_planilla = {}
            for row in range(2, ws_rcu.max_row + 1):
                cell_val = ws_rcu.cell(row=row, column=1).value
                if cell_val: ids_planilla[str(cell_val).strip().upper()] = row

            dict_sla_marcas = {}
            if "SLA" in wb.sheetnames:
                ws_sla = wb["SLA"]
                for row in range(2, ws_sla.max_row + 1):
                    id_sla = ws_sla.cell(row=row, column=1).value
                    sla_marca = ws_sla.cell(row=row, column=9).value
                    if id_sla and sla_marca: dict_sla_marcas[str(id_sla).strip().upper()] = str(sla_marca)

            actualizados = 0
            nuevos = 0
            
            for _, row in df_nuevo.iterrows():
                id_raw = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
                if not id_raw or id_raw.upper() in ["ID", "ID2", "ID3", "ID4"]: continue
                id_upper = id_raw.upper()
                
                # Extract fields
                address2 = str(row.get('ADDRESS2', '')) if pd.notna(row.get('ADDRESS2', None)) else ""
                region = str(row.get('REGION', '')) if pd.notna(row.get('REGION', None)) else ""
                zone = str(row.get('ZONE', '')) if pd.notna(row.get('ZONE', None)) else ""
                
                custodio_norm = self.normalizar_custodio(region, zone)
                sla_val = dict_sla_marcas.get(id_upper, "Sin SLA")

                if id_upper in ids_planilla:
                    # Update RCU (simplified for this migration)
                    ws_rcu.cell(row=ids_planilla[id_upper], column=3, value=address2)
                    ws_rcu.cell(row=ids_planilla[id_upper], column=12, value=sla_val)
                    actualizados += 1
                else:
                    # Add to RCU
                    new_row = ws_rcu.max_row + 1
                    ws_rcu.cell(row=new_row, column=1, value=id_upper)
                    ws_rcu.cell(row=new_row, column=3, value=address2)
                    ws_rcu.cell(row=new_row, column=12, value=sla_val)
                    nuevos += 1
                    ids_planilla[id_upper] = new_row

                # Update/Add to UNIFICADO
                if id_upper in ids_unif:
                    u_row = ids_unif[id_upper]
                else:
                    u_row = ws_unif.max_row + 1
                    ws_unif.cell(row=u_row, column=1, value=id_upper)
                
                ws_unif.cell(row=u_row, column=2, value=address2)
                ws_unif.cell(row=u_row, column=3, value=custodio_norm)
                ws_unif.cell(row=u_row, column=4, value=sla_val)
                ws_unif.cell(row=u_row, column=12, value=sla_val)

            wb.save(self.excel_path)
            return True, {
                'actualizados': actualizados,
                'nuevos': nuevos,
                'total_procesados': actualizados + nuevos,
                'total_archivo': len(df_nuevo)
            }
        except Exception as e:
            return False, str(e)
