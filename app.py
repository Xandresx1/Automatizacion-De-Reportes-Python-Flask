"""
Aplicación principal de la Herramienta de Automatización de Reportes.

Esta aplicación Flask proporciona una interfaz web amigable para:
- Subir archivos de datos (CSV/Excel)
- Procesar y limpiar datos automáticamente
- Generar reportes con gráficos
- Enviar reportes por email
"""

from flask import Flask, render_template, request, redirect, url_for, flash, send_file, jsonify
from werkzeug.utils import secure_filename
import os
import logging
from datetime import datetime
from dotenv import load_dotenv
import json

# Importar módulos personalizados
from data_processor import DataProcessor
from report_generator import ReportGenerator
from email_sender import EmailSender

# Cargar variables de entorno
load_dotenv()

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Crear aplicación Flask
app = Flask(__name__)
app.secret_key = 'tu_clave_secreta_aqui_cambiala_en_produccion'

# Configuración de archivos
UPLOAD_FOLDER = 'uploads'
REPORTS_FOLDER = 'reports'
ALLOWED_EXTENSIONS = {'csv', 'xlsx', 'xls'}

# Crear carpetas si no existen
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['REPORTS_FOLDER'] = REPORTS_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max

# Variables globales para mantener estado
current_data_processor = None
current_report_generator = None
email_sender = EmailSender()

def allowed_file(filename):
    """Verifica si el archivo tiene una extensión permitida."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_file_size_mb(file_path):
    """Obtiene el tamaño del archivo en MB."""
    try:
        size_bytes = os.path.getsize(file_path)
        return round(size_bytes / (1024 * 1024), 2)
    except:
        return 0

@app.route('/')
def index():
    """Página principal con formulario de carga de archivos."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Maneja la carga de archivos y el procesamiento inicial."""
    global current_data_processor, current_report_generator
    
    try:
        # Verificar si se subió un archivo
        if 'file' not in request.files:
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)
        
        file = request.files['file']
        
        # Verificar si el archivo tiene nombre
        if file.filename == '':
            flash('No se seleccionó ningún archivo', 'error')
            return redirect(request.url)
        
        # Verificar extensión del archivo
        if not allowed_file(file.filename):
            flash('Tipo de archivo no permitido. Usa CSV o Excel (.xlsx, .xls)', 'error')
            return redirect(request.url)
        
        # Guardar archivo
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename_with_timestamp = f"{timestamp}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename_with_timestamp)
        file.save(file_path)
        
        # Inicializar procesador de datos
        current_data_processor = DataProcessor()
        
        # Cargar datos
        success, message = current_data_processor.load_data(file_path)
        if not success:
            flash(f'Error al cargar datos: {message}', 'error')
            return redirect(url_for('index'))
        
        # Limpiar datos
        success, message = current_data_processor.clean_data()
        if not success:
            flash(f'Error al limpiar datos: {message}', 'error')
            return redirect(url_for('index'))
        
        # Obtener resumen de datos
        data_summary = current_data_processor.get_data_summary()
        sample_data = current_data_processor.get_sample_data(5)
        
        # Inicializar generador de reportes
        current_report_generator = ReportGenerator()
        current_report_generator.set_data(current_data_processor.get_cleaned_data())
        
        flash(f'Archivo cargado exitosamente: {filename}', 'success')
        
        return render_template('data_preview.html', 
                             data_summary=data_summary,
                             sample_data=sample_data.to_html(classes='table table-striped', index=False),
                             filename=filename)
    
    except Exception as e:
        logger.error(f"Error en upload_file: {str(e)}")
        flash(f'Error inesperado: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/generate_report', methods=['POST'])
def generate_report():
    """Genera el reporte basado en los datos procesados."""
    global current_report_generator
    
    try:
        if current_report_generator is None:
            flash('No hay datos procesados para generar reporte', 'error')
            return redirect(url_for('index'))
        
        # Obtener parámetros del formulario
        report_title = request.form.get('report_title', 'Reporte Automático')
        report_description = request.form.get('report_description', '')
        recipients = request.form.get('recipients', '').split(',')
        recipients = [email.strip() for email in recipients if email.strip()]
        
        # Configurar información del reporte
        current_report_generator.set_report_info(
            title=report_title,
            description=report_description,
            recipients=recipients
        )
        
        # Generar reporte Excel
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_filename = f"reporte_{timestamp}.xlsx"
        excel_path = os.path.join(app.config['REPORTS_FOLDER'], excel_filename)
        
        success, message = current_report_generator.generate_excel_report(excel_path)
        if not success:
            flash(f'Error al generar reporte: {message}', 'error')
            return redirect(url_for('index'))
        
        # Obtener resumen del reporte
        report_summary = current_report_generator.get_report_summary()
        
        flash(f'Reporte generado exitosamente: {excel_filename}', 'success')
        
        return render_template('report_generated.html',
                             report_summary=report_summary,
                             excel_filename=excel_filename,
                             recipients=recipients)
    
    except Exception as e:
        logger.error(f"Error en generate_report: {str(e)}")
        flash(f'Error al generar reporte: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/download_report/<filename>')
def download_report(filename):
    """Permite descargar el reporte generado."""
    try:
        file_path = os.path.join(app.config['REPORTS_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True)
        else:
            flash('Archivo no encontrado', 'error')
            return redirect(url_for('index'))
    except Exception as e:
        flash(f'Error al descargar archivo: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/send_email', methods=['POST'])
def send_email():
    """Envía el reporte por email."""
    global current_report_generator, email_sender
    
    try:
        # Configurar email desde variables de entorno
        success, message = email_sender.configure_from_env()
        if not success:
            flash(f'Error de configuración de email: {message}', 'error')
            return redirect(url_for('index'))
        
        # Obtener parámetros
        recipients = request.form.get('recipients', '').split(',')
        recipients = [email.strip() for email in recipients if email.strip()]
        
        if not recipients:
            flash('No se especificaron destinatarios', 'error')
            return redirect(url_for('index'))
        
        # Obtener información del reporte
        report_summary = current_report_generator.get_report_summary()
        data_summary = current_data_processor.get_data_summary()
        
        # Crear cuerpo del email
        email_body = email_sender.create_email_body(
            report_title=report_summary['title'],
            data_summary=data_summary,
            charts_created=report_summary.get('charts_list', [])
        )
        
        # Enviar email
        subject = f"📊 {report_summary['title']} - {datetime.now().strftime('%d/%m/%Y')}"
        
        # Buscar archivo Excel para adjuntar
        excel_filename = request.form.get('excel_filename')
        attachment_path = None
        if excel_filename:
            attachment_path = os.path.join(app.config['REPORTS_FOLDER'], excel_filename)
        
        success, message = email_sender.send_report(
            recipients=recipients,
            subject=subject,
            body=email_body,
            attachment_path=attachment_path
        )
        
        if success:
            flash(f'Email enviado exitosamente a {len(recipients)} destinatarios', 'success')
        else:
            flash(f'Error al enviar email: {message}', 'error')
        
        return redirect(url_for('index'))
    
    except Exception as e:
        logger.error(f"Error en send_email: {str(e)}")
        flash(f'Error al enviar email: {str(e)}', 'error')
        return redirect(url_for('index'))

@app.route('/test_email_connection')
def test_email_connection():
    """Prueba la conexión de email."""
    try:
        # Configurar email
        success, message = email_sender.configure_from_env()
        if not success:
            return jsonify({'success': False, 'message': message})
        
        # Probar conexión
        success, message = email_sender.test_connection()
        return jsonify({'success': success, 'message': message})
    
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

@app.route('/api/data_summary')
def api_data_summary():
    """API para obtener resumen de datos."""
    global current_data_processor
    
    if current_data_processor is None:
        return jsonify({'error': 'No hay datos procesados'})
    
    try:
        summary = current_data_processor.get_data_summary()
        return jsonify(summary)
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/api/sample_data')
def api_sample_data():
    """API para obtener muestra de datos."""
    global current_data_processor
    
    if current_data_processor is None:
        return jsonify({'error': 'No hay datos procesados'})
    
    try:
        sample_data = current_data_processor.get_sample_data(10)
        return jsonify({
            'columns': list(sample_data.columns),
            'data': sample_data.to_dict('records')
        })
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/help')
def help_page():
    """Página de ayuda y documentación."""
    return render_template('help.html')

@app.route('/about')
def about_page():
    """Página sobre la herramienta."""
    return render_template('about.html')

@app.errorhandler(413)
def too_large(e):
    """Maneja archivos demasiado grandes."""
    flash('El archivo es demasiado grande. Máximo 16MB.', 'error')
    return redirect(url_for('index'))

@app.errorhandler(404)
def not_found(e):
    """Maneja páginas no encontradas."""
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_error(e):
    """Maneja errores internos del servidor."""
    flash('Error interno del servidor. Intenta de nuevo.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    # Verificar configuración de email al iniciar
    try:
        success, message = email_sender.configure_from_env()
        if success:
            logger.info("Configuración de email cargada exitosamente")
        else:
            logger.warning(f"Configuración de email no disponible: {message}")
    except Exception as e:
        logger.warning(f"No se pudo cargar configuración de email: {e}")
    
    # Iniciar aplicación
    logger.info("Iniciando aplicación de automatización de reportes...")
    app.run(debug=True, host='0.0.0.0', port=5000)
