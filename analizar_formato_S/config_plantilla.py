"""
Script de Configuración - Plantilla de Tesis
Este script solicita un documento Word local y lo copia a la carpeta plantillas/
"""

import os
import shutil
from pathlib import Path

# Colores para terminal (opcional, funciona en Windows 10+)
class Color:
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BLUE = '\033[94m'
    END = '\033[0m'

def crear_estructura_carpetas():
    """Crea la estructura de carpetas del proyecto"""
    carpetas = ['plantillas', 'generados']
    
    for carpeta in carpetas:
        if not os.path.exists(carpeta):
            os.makedirs(carpeta)
            print(f"{Color.GREEN}✓{Color.END} Carpeta '{carpeta}' creada")
        else:
            print(f"{Color.BLUE}ℹ{Color.END} Carpeta '{carpeta}' ya existe")
    
    # Crear .gitignore en generados para no versionar documentos generados
    gitignore_path = os.path.join('generados', '.gitignore')
    if not os.path.exists(gitignore_path):
        with open(gitignore_path, 'w') as f:
            f.write("# Ignorar todos los archivos generados\n*.docx\n*.doc\n")
        print(f"{Color.GREEN}✓{Color.END} .gitignore creado en 'generados/'")

def validar_documento(ruta):
    """Valida que el archivo sea un documento Word válido"""
    if not os.path.exists(ruta):
        print(f"{Color.RED}✗{Color.END} El archivo no existe: {ruta}")
        return False
    
    if not ruta.lower().endswith(('.docx', '.doc')):
        print(f"{Color.RED}✗{Color.END} El archivo debe ser .docx o .doc")
        return False
    
    # Verificar que el archivo no esté vacío
    if os.path.getsize(ruta) == 0:
        print(f"{Color.RED}✗{Color.END} El archivo está vacío")
        return False
    
    return True

def copiar_plantilla(ruta_origen):
    """Copia el documento a la carpeta plantillas/"""
    destino = os.path.join('plantillas', 'formato_tesis.docx')
    
    # Si ya existe una plantilla, preguntar si reemplazar
    if os.path.exists(destino):
        print(f"\n{Color.YELLOW}⚠{Color.END} Ya existe una plantilla en: {destino}")
        respuesta = input("¿Deseas reemplazarla? (s/n): ").lower().strip()
        
        if respuesta != 's':
            print(f"{Color.BLUE}ℹ{Color.END} Operación cancelada. Se mantendrá la plantilla actual.")
            return False
    
    try:
        shutil.copy2(ruta_origen, destino)
        print(f"{Color.GREEN}✓{Color.END} Plantilla copiada exitosamente a: {destino}")
        
        # Mostrar información del archivo
        tamaño = os.path.getsize(destino) / 1024  # KB
        print(f"{Color.BLUE}ℹ{Color.END} Tamaño del archivo: {tamaño:.2f} KB")
        
        return True
    except Exception as e:
        print(f"{Color.RED}✗{Color.END} Error al copiar el archivo: {e}")
        return False

def main():
    print("="*60)
    print(f"{Color.BLUE}CONFIGURACIÓN DE PLANTILLA DE TESIS{Color.END}")
    print("="*60)
    print()
    
    # Crear estructura de carpetas
    print("📁 Verificando estructura de carpetas...")
    crear_estructura_carpetas()
    print()
    
    # Solicitar ruta del documento
    print("📄 Ingresa la ruta del documento Word que usarás como plantilla:")
    print(f"{Color.YELLOW}Ejemplo:{Color.END} C:/Users/Steeve/Desktop/mi_tesis.docx")
    print(f"{Color.YELLOW}O simplemente arrastra el archivo aquí{Color.END}")
    print()
    
    ruta_documento = input("Ruta del documento: ").strip().strip('"').strip("'")
    
    # Convertir ruta relativa a absoluta si es necesario
    ruta_documento = os.path.abspath(ruta_documento)
    
    print()
    print(f"📍 Ruta detectada: {ruta_documento}")
    print()
    
    # Validar documento
    print("🔍 Validando documento...")
    if not validar_documento(ruta_documento):
        print(f"\n{Color.RED}✗{Color.END} La configuración falló. Por favor, verifica el archivo.")
        return
    
    print(f"{Color.GREEN}✓{Color.END} Documento válido")
    print()
    
    # Copiar plantilla
    print("📋 Copiando plantilla...")
    if copiar_plantilla(ruta_documento):
        print()
        print("="*60)
        print(f"{Color.GREEN}✅ CONFIGURACIÓN COMPLETADA EXITOSAMENTE{Color.END}")
        print("="*60)
        print()
        print("Próximos pasos:")
        print(f"  1. Ejecuta {Color.BLUE}analizar_formato.py{Color.END} para analizar la estructura")
        print(f"  2. Usa {Color.BLUE}generar_tesis.py{Color.END} para crear nuevos documentos")
        print()
    else:
        print(f"\n{Color.RED}✗{Color.END} La configuración falló.")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(f"\n\n{Color.YELLOW}⚠{Color.END} Operación cancelada por el usuario.")
    except Exception as e:
        print(f"\n{Color.RED}✗{Color.END} Error inesperado: {e}")