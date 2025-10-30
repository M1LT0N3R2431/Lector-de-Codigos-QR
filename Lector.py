# qr_scanner_auto_fixed.py - Versión corregida para PyInstaller
import cv2
import pandas as pd
from pyzbar.pyzbar import decode
from datetime import datetime
import time
import sys
import os

def resource_path(relative_path):
    """Obtiene la ruta absoluta al recurso, funciona para dev y para PyInstaller"""
    try:
        # PyInstaller crea una carpeta temporal y almacena la ruta en _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def escanear_qr_automatico():
    """Escanea QR automáticamente sin intervención del usuario"""
    
    print("🔍 Iniciando escáner QR automático...")
    print("Enfoca los códigos QR con la cámara USB")
    print("Los datos se guardan automáticamente en 'qr_escaneados.xlsx'")
    print("Presiona CTRL+C en la terminal para detener")
    print("O presiona 'q' en la ventana de la cámara para salir")
    
    datos = []
    
    # Intentar diferentes índices de cámara
    cap = None
    for i in range(3):  # Probar cámaras 0, 1, 2
        cap = cv2.VideoCapture(i)
        if cap.isOpened():
            print(f"✅ Cámara {i} inicializada correctamente")
            break
        else:
            print(f"❌ No se pudo abrir cámara {i}")
            cap = None
    
    if cap is None:
        print("❌ No se pudo inicializar ninguna cámara")
        input("Presiona Enter para salir...")
        return
    
    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                print("⚠️ No se pudo leer el frame de la cámara")
                time.sleep(0.1)
                continue
            
            # Detectar QR
            qr_codes = decode(frame)
            
            for qr in qr_codes:
                try:
                    contenido = qr.data.decode('utf-8')
                    
                    # Evitar duplicados
                    if contenido not in [d['contenido'] for d in datos]:
                        nuevo_dato = {
                            'numero': len(datos) + 1,
                            'contenido': contenido,
                            'tipo': qr.type,
                            'fecha_hora': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        datos.append(nuevo_dato)
                        
                        print(f"✅ [{len(datos)}] {contenido}")
                        
                        # Guardar automáticamente en Excel
                        guardar_excel(datos)
                except Exception as e:
                    print(f"❌ Error decodificando QR: {e}")
            
            # Mostrar cámara
            cv2.imshow('Escáner QR - Presiona Q para salir', frame)
            
            # Salir con Q
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
                
    except KeyboardInterrupt:
        print("\n🛑 Programa interrumpido por el usuario (CTRL+C)")
    except Exception as e:
        print(f"\n❌ Error inesperado: {e}")
    finally:
        if cap and cap.isOpened():
            cap.release()
        cv2.destroyAllWindows()
        
        if datos:
            print(f"\n📊 Resumen: {len(datos)} códigos QR escaneados")
            print(f"💾 Datos guardados en: qr_escaneados.xlsx")
        else:
            print("\n📊 No se escanearon códigos QR")
        
        input("\nPresiona Enter para cerrar...")

def guardar_excel(datos):
    """Guarda los datos en Excel automáticamente"""
    try:
        df = pd.DataFrame(datos)
        df.to_excel('qr_escaneados.xlsx', index=False)
    except Exception as e:
        print(f"❌ Error guardando Excel: {e}")

if __name__ == "__main__":
    escanear_qr_automatico()