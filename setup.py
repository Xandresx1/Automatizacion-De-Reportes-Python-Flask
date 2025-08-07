#!/usr/bin/env python3
"""
Script de configuración para la Herramienta de Automatización de Reportes.

Este script ayuda a configurar el entorno y las dependencias necesarias
para ejecutar la aplicación.
"""

import os
import sys
import subprocess
import platform

def print_banner():
    """Imprime el banner de bienvenida."""
    print("=" * 60)
    print("🚀 Herramienta de Automatización de Reportes")
    print("=" * 60)
    print()

def check_python_version():
    """Verifica la versión de Python."""
    print("📋 Verificando versión de Python...")
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print("❌ Error: Se requiere Python 3.8 o superior")
        print(f"   Versión actual: {version.major}.{version.minor}.{version.micro}")
        return False
    else:
        print(f"✅ Python {version.major}.{version.minor}.{version.micro} - OK")
        return True

def create_virtual_environment():
    """Crea el entorno virtual."""
    print("\n🐍 Creando entorno virtual...")
    
    if os.path.exists("venv"):
        print("✅ Entorno virtual ya existe")
        return True
    
    try:
        subprocess.run([sys.executable, "-m", "venv", "venv"], check=True)
        print("✅ Entorno virtual creado exitosamente")
        return True
    except subprocess.CalledProcessError:
        print("❌ Error al crear el entorno virtual")
        return False

def install_dependencies():
    """Instala las dependencias."""
    print("\n📦 Instalando dependencias...")
    
    # Determinar el comando de pip según el sistema operativo
    if platform.system() == "Windows":
        pip_cmd = "venv\\Scripts\\pip"
    else:
        pip_cmd = "venv/bin/pip"
    
    try:
        subprocess.run([pip_cmd, "install", "-r", "requirements.txt"], check=True)
        print("✅ Dependencias instaladas exitosamente")
        return True
    except subprocess.CalledProcessError:
        print("❌ Error al instalar dependencias")
        return False

def create_env_file():
    """Crea el archivo .env si no existe."""
    print("\n⚙️  Configurando variables de entorno...")
    
    if os.path.exists(".env"):
        print("✅ Archivo .env ya existe")
        return True
    
    env_content = """# Configuración de Email
# Reemplaza estos valores con tus credenciales reales

EMAIL_USER=tu_correo@gmail.com
EMAIL_PASSWORD=tu_contraseña_de_aplicacion
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# Notas importantes:
# 1. Para Gmail, necesitas crear una "contraseña de aplicación"
# 2. Ve a Configuración de tu cuenta Google → Seguridad → Contraseñas de aplicación
# 3. Crea una nueva contraseña de aplicación y úsala aquí
# 4. NUNCA uses tu contraseña normal de Gmail
"""
    
    try:
        with open(".env", "w", encoding="utf-8") as f:
            f.write(env_content)
        print("✅ Archivo .env creado")
        print("⚠️  IMPORTANTE: Edita el archivo .env con tus credenciales reales")
        return True
    except Exception as e:
        print(f"❌ Error al crear archivo .env: {e}")
        return False

def create_directories():
    """Crea las carpetas necesarias."""
    print("\n📁 Creando carpetas necesarias...")
    
    directories = ["uploads", "reports", "static"]
    
    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            print(f"✅ Carpeta '{directory}' creada")
        else:
            print(f"✅ Carpeta '{directory}' ya existe")

def print_activation_instructions():
    """Imprime las instrucciones de activación."""
    print("\n" + "=" * 60)
    print("🎉 ¡Configuración completada!")
    print("=" * 60)
    
    print("\n📋 Para activar el entorno virtual:")
    
    if platform.system() == "Windows":
        print("   venv\\Scripts\\activate")
    else:
        print("   source venv/bin/activate")
    
    print("\n🚀 Para ejecutar la aplicación:")
    print("   python app.py")
    
    print("\n🌐 La aplicación estará disponible en:")
    print("   http://localhost:5000")
    
    print("\n⚠️  IMPORTANTE:")
    print("   1. Edita el archivo .env con tus credenciales de email")
    print("   2. Para Gmail, usa una contraseña de aplicación")
    print("   3. Revisa el README.md para más información")
    
    print("\n📚 Archivos útiles:")
    print("   - README.md: Documentación completa")
    print("   - ejemplo_datos.csv: Datos de ejemplo para probar")
    print("   - config_ejemplo.txt: Ejemplo de configuración")

def main():
    """Función principal del script."""
    print_banner()
    
    # Verificar Python
    if not check_python_version():
        return
    
    # Crear entorno virtual
    if not create_virtual_environment():
        return
    
    # Instalar dependencias
    if not install_dependencies():
        return
    
    # Crear archivo .env
    create_env_file()
    
    # Crear carpetas
    create_directories()
    
    # Mostrar instrucciones
    print_activation_instructions()

if __name__ == "__main__":
    main()
