import time

try:
    import win32com.client
    import pythoncom
    WINDOWS_OUTLOOK_AVAILABLE = True
except ImportError:
    WINDOWS_OUTLOOK_AVAILABLE = False

def crear_correo_outlook(to, cc, asunto, cuerpo):
    if not WINDOWS_OUTLOOK_AVAILABLE:
        return False, "La integración con Outlook requiere Windows y la librería pywin32."
    
    try:
        # Inicializar COM para el hilo actual (Crítico para Flask)
        pythoncom.CoInitialize()
        
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        
        mail.To = to
        mail.CC = cc
        mail.Subject = asunto
        
        # Primero mostramos el correo para que cargue la firma predeterminada
        mail.Display() 
        
        # Esperamos un momento para que Outlook genere la ventana
        time.sleep(1)
        
        # Insertamos el cuerpo HTML respetando la firma de Outlook
        mail.HTMLBody = cuerpo + mail.HTMLBody
        
        return True, "Correo mostrado en Outlook correctamente."
    except Exception as e:
        return False, f"Error al conectar con Outlook: {str(e)}"
    finally:
        # Liberamos los recursos COM
        pythoncom.CoUninitialize()
