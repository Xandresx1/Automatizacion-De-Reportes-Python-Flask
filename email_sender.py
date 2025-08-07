"""
Módulo de envío automático de emails para la herramienta de automatización de reportes.

Este módulo maneja el envío seguro de reportes por correo electrónico,
incluyendo configuración de SMTP, manejo de errores y validación de destinatarios.
"""

import smtplib
import os
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from typing import List, Tuple, Optional
from datetime import datetime
import re

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class EmailSender:
    """
    Clase para el envío automático de reportes por email.
    
    Esta clase maneja toda la configuración de SMTP, validación de emails
    y envío seguro de reportes con archivos adjuntos.
    """
    
    def __init__(self):
        """Inicializa el remitente de emails."""
        self.smtp_server = None
        self.smtp_port = None
        self.email_user = None
        self.email_password = None
        self.is_configured = False
        
    def configure_from_env(self) -> Tuple[bool, str]:
        """
        Configura el remitente desde variables de entorno.
        
        Esta función lee la configuración de email desde variables de entorno,
        lo cual es más seguro que hardcodear credenciales en el código.
        
        Returns:
            Tuple con (éxito, mensaje)
        """
        try:
            # Cargar variables de entorno
            self.email_user = os.getenv('EMAIL_USER')
            self.email_password = os.getenv('EMAIL_PASSWORD')
            self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.gmail.com')
            self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
            
            # Validar que las credenciales estén presentes
            if not self.email_user or not self.email_password:
                return False, "Credenciales de email no configuradas. Verifica EMAIL_USER y EMAIL_PASSWORD en el archivo .env"
            
            # Validar formato del email
            if not self._is_valid_email(self.email_user):
                return False, "Formato de email inválido en EMAIL_USER"
            
            self.is_configured = True
            logger.info(f"Configuración de email exitosa para: {self.email_user}")
            return True, "Configuración de email exitosa"
            
        except Exception as e:
            logger.error(f"Error al configurar email: {str(e)}")
            return False, f"Error al configurar email: {str(e)}"
    
    def configure_manually(self, email_user: str, email_password: str, 
                          smtp_server: str = 'smtp.gmail.com', smtp_port: int = 587) -> Tuple[bool, str]:
        """
        Configura el remitente manualmente.
        
        Args:
            email_user: Email del remitente
            email_password: Contraseña o contraseña de aplicación
            smtp_server: Servidor SMTP
            smtp_port: Puerto SMTP
            
        Returns:
            Tuple con (éxito, mensaje)
        """
        try:
            # Validar formato del email
            if not self._is_valid_email(email_user):
                return False, "Formato de email inválido"
            
            self.email_user = email_user
            self.email_password = email_password
            self.smtp_server = smtp_server
            self.smtp_port = smtp_port
            self.is_configured = True
            
            logger.info(f"Configuración manual de email exitosa para: {email_user}")
            return True, "Configuración manual de email exitosa"
            
        except Exception as e:
            logger.error(f"Error en configuración manual: {str(e)}")
            return False, f"Error en configuración manual: {str(e)}"
    
    def send_report(self, recipients: List[str], subject: str, body: str, 
                   attachment_path: str = None) -> Tuple[bool, str]:
        """
        Envía un reporte por email.
        
        Args:
            recipients: Lista de destinatarios
            subject: Asunto del email
            body: Cuerpo del email
            attachment_path: Ruta al archivo adjunto (opcional)
            
        Returns:
            Tuple con (éxito, mensaje)
        """
        if not self.is_configured:
            return False, "Email no configurado. Usa configure_from_env() o configure_manually() primero."
        
        # Validar destinatarios
        valid_recipients = []
        invalid_emails = []
        
        for email in recipients:
            if self._is_valid_email(email):
                valid_recipients.append(email)
            else:
                invalid_emails.append(email)
        
        if not valid_recipients:
            return False, "No hay emails válidos en la lista de destinatarios"
        
        if invalid_emails:
            logger.warning(f"Emails inválidos ignorados: {invalid_emails}")
        
        try:
            # Crear mensaje
            msg = MIMEMultipart()
            msg['From'] = self.email_user
            msg['To'] = ', '.join(valid_recipients)
            msg['Subject'] = subject
            
            # Agregar cuerpo del email
            msg.attach(MIMEText(body, 'html'))
            
            # Agregar archivo adjunto si existe
            if attachment_path and os.path.exists(attachment_path):
                self._add_attachment(msg, attachment_path)
                logger.info(f"Archivo adjunto agregado: {attachment_path}")
            
            # Enviar email
            self._send_email(msg, valid_recipients)
            
            return True, f"Email enviado exitosamente a {len(valid_recipients)} destinatarios"
            
        except Exception as e:
            logger.error(f"Error al enviar email: {str(e)}")
            return False, f"Error al enviar email: {str(e)}"
    
    def _send_email(self, msg: MIMEMultipart, recipients: List[str]):
        """
        Envía el email usando SMTP.
        
        Args:
            msg: Mensaje MIME a enviar
            recipients: Lista de destinatarios
        """
        try:
            # Conectar al servidor SMTP
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()  # Habilitar TLS para seguridad
            
            # Autenticarse
            server.login(self.email_user, self.email_password)
            
            # Enviar email
            text = msg.as_string()
            server.sendmail(self.email_user, recipients, text)
            
            # Cerrar conexión
            server.quit()
            
            logger.info(f"Email enviado exitosamente a {len(recipients)} destinatarios")
            
        except smtplib.SMTPAuthenticationError:
            raise Exception("Error de autenticación. Verifica tu email y contraseña. Para Gmail, usa una contraseña de aplicación.")
        except smtplib.SMTPRecipientsRefused:
            raise Exception("Uno o más destinatarios fueron rechazados por el servidor.")
        except smtplib.SMTPServerDisconnected:
            raise Exception("Conexión con el servidor SMTP perdida.")
        except Exception as e:
            raise Exception(f"Error de SMTP: {str(e)}")
    
    def _add_attachment(self, msg: MIMEMultipart, file_path: str):
        """
        Agrega un archivo adjunto al mensaje.
        
        Args:
            msg: Mensaje MIME
            file_path: Ruta al archivo a adjuntar
        """
        try:
            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            # Codificar el archivo
            encoders.encode_base64(part)
            
            # Agregar headers
            filename = os.path.basename(file_path)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {filename}'
            )
            
            msg.attach(part)
            
        except Exception as e:
            logger.error(f"Error al adjuntar archivo {file_path}: {e}")
            raise
    
    def _is_valid_email(self, email: str) -> bool:
        """
        Valida el formato de un email.
        
        Args:
            email: Email a validar
            
        Returns:
            True si el email tiene formato válido
        """
        if not email or '@' not in email:
            return False
        
        # Patrón básico para validar email
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def create_email_body(self, report_title: str, data_summary: dict, 
                         charts_created: List[str] = None) -> str:
        """
        Crea el cuerpo del email con información del reporte.
        
        Args:
            report_title: Título del reporte
            data_summary: Resumen de los datos procesados
            charts_created: Lista de gráficos creados
            
        Returns:
            Cuerpo del email en formato HTML
        """
        html_body = f"""
        <html>
        <head>
            <style>
                body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                .header {{ background-color: #f8f9fa; padding: 20px; border-radius: 5px; }}
                .content {{ padding: 20px; }}
                .summary {{ background-color: #e9ecef; padding: 15px; border-radius: 5px; margin: 10px 0; }}
                .footer {{ background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin-top: 20px; font-size: 12px; color: #666; }}
                .highlight {{ color: #007bff; font-weight: bold; }}
                .success {{ color: #28a745; }}
            </style>
        </head>
        <body>
            <div class="header">
                <h2>📊 {report_title}</h2>
                <p>Reporte generado automáticamente el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}</p>
            </div>
            
            <div class="content">
                <h3>📈 Resumen del Procesamiento</h3>
                <div class="summary">
                    <p><span class="highlight">Registros procesados:</span> {data_summary.get('filas_finales', 0):,}</p>
                    <p><span class="highlight">Registros originales:</span> {data_summary.get('filas_originales', 0):,}</p>
                    <p><span class="highlight">Registros eliminados:</span> {data_summary.get('filas_eliminadas', 0):,}</p>
                    <p><span class="highlight">Columnas analizadas:</span> {data_summary.get('columnas_totales', 0)}</p>
                    <p><span class="highlight">Columnas numéricas:</span> {data_summary.get('columnas_numericas', 0)}</p>
                    <p><span class="highlight">Columnas de fecha:</span> {data_summary.get('columnas_fecha', 0)}</p>
                </div>
                
                {self._create_charts_section(charts_created) if charts_created else ''}
                
                <h3>📋 Contenido del Reporte</h3>
                <ul>
                    <li><strong>Resumen Ejecutivo:</strong> Estadísticas principales y hallazgos clave</li>
                    <li><strong>Datos Procesados:</strong> Dataset completo con limpieza aplicada</li>
                    <li><strong>Gráficos Automáticos:</strong> Visualizaciones generadas automáticamente</li>
                    <li><strong>Análisis Estadístico:</strong> Estadísticas descriptivas detalladas</li>
                </ul>
            </div>
            
            <div class="footer">
                <p>🤖 Este reporte fue generado automáticamente por la Herramienta de Automatización de Reportes</p>
                <p>💡 <em>Recuerda revisar la carpeta de spam si el reporte no llega al correo principal.</em></p>
                <p>📧 Si tienes problemas con la recepción, verifica que tu proveedor de email no esté bloqueando archivos adjuntos.</p>
            </div>
        </body>
        </html>
        """
        
        return html_body
    
    def _create_charts_section(self, charts_created: List[str]) -> str:
        """
        Crea la sección de gráficos para el email.
        
        Args:
            charts_created: Lista de gráficos creados
            
        Returns:
            HTML de la sección de gráficos
        """
        if not charts_created:
            return ""
        
        charts_html = "<h3>📊 Gráficos Generados</h3><ul>"
        for chart in charts_created:
            charts_html += f"<li><span class='success'>✓</span> {chart}</li>"
        charts_html += "</ul>"
        
        return charts_html
    
    def test_connection(self) -> Tuple[bool, str]:
        """
        Prueba la conexión SMTP sin enviar emails.
        
        Returns:
            Tuple con (éxito, mensaje)
        """
        if not self.is_configured:
            return False, "Email no configurado"
        
        try:
            # Conectar al servidor SMTP
            server = smtplib.SMTP(self.smtp_server, self.smtp_port)
            server.starttls()
            
            # Intentar autenticarse
            server.login(self.email_user, self.email_password)
            
            # Cerrar conexión
            server.quit()
            
            return True, "Conexión SMTP exitosa"
            
        except smtplib.SMTPAuthenticationError:
            return False, "Error de autenticación. Verifica tu email y contraseña. Para Gmail, usa una contraseña de aplicación."
        except Exception as e:
            return False, f"Error de conexión: {str(e)}"
    
    def get_configuration_info(self) -> dict:
        """
        Obtiene información de la configuración actual.
        
        Returns:
            Diccionario con información de configuración
        """
        return {
            'is_configured': self.is_configured,
            'email_user': self.email_user,
            'smtp_server': self.smtp_server,
            'smtp_port': self.smtp_port,
            'has_password': bool(self.email_password)
        }
