import json
import os
import ezodf
import smtplib
import time
import threading
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from tkinter import filedialog, messagebox, scrolledtext
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from PIL import Image, ImageTk

CONFIG_FILE = 'config.json'


def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def save_config(config):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)


class EmailSenderApp(ttk.Window):
    def __init__(self):
        super().__init__(themename='flatly')
        self.title('Enviador de Correos')
        self.geometry('550x800')  # Increased height for logo and console

        self.config = load_config()
        
        # Logo
        try:
            logo_path = 'umbreon.png'
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                # Resize if needed
                logo_img = logo_img.resize((100, 100), Image.LANCZOS)
                logo_photo = ImageTk.PhotoImage(logo_img)
                logo_label = ttk.Label(self, image=logo_photo)
                logo_label.image = logo_photo  # Keep a reference
                logo_label.pack(pady=(10, 5))
            else:
                ttk.Label(self, text="Logo no encontrado (umbreon.png)").pack(pady=(10, 5))
        except Exception as e:
            print(f"Error loading logo: {e}")

        # SMTP Server
        ttk.Label(self, text='Servidor SMTP').pack(pady=(10, 0))
        self.smtp_entry = ttk.Entry(self)
        self.smtp_entry.pack(fill=X, padx=20)
        self.smtp_entry.insert(0, self.config.get('smtp_server', ''))

        # Port
        ttk.Label(self, text='Puerto').pack(pady=(10, 0))
        self.port_entry = ttk.Entry(self)
        self.port_entry.pack(fill=X, padx=20)
        self.port_entry.insert(0, self.config.get('port', '587'))

        # User Email
        ttk.Label(self, text='Correo Electrónico').pack(pady=(10, 0))
        self.user_entry = ttk.Entry(self)
        self.user_entry.pack(fill=X, padx=20)
        self.user_entry.insert(0, self.config.get('user', ''))

        # Password
        ttk.Label(self, text='Contraseña').pack(pady=(10, 0))
        self.pw_entry = ttk.Entry(self, show='*')
        self.pw_entry.pack(fill=X, padx=20)
        self.pw_entry.insert(0, self.config.get('password', ''))

        # HTML Template
        ttk.Label(self, text='Plantilla HTML').pack(pady=(10, 0))
        frame_html = ttk.Frame(self)
        frame_html.pack(fill=X, padx=20)
        self.html_path = ttk.StringVar()
        ttk.Entry(frame_html, textvariable=self.html_path).pack(side=LEFT, fill=X, expand=YES)
        ttk.Button(frame_html, text='Examinar', bootstyle=PRIMARY, command=self.browse_html).pack(side=LEFT, padx=5)

        # ODS File
        ttk.Label(self, text='Archivo ODS').pack(pady=(10, 0))
        frame_ods = ttk.Frame(self)
        frame_ods.pack(fill=X, padx=20)
        self.ods_path = ttk.StringVar()
        ttk.Entry(frame_ods, textvariable=self.ods_path).pack(side=LEFT, fill=X, expand=YES)
        ttk.Button(frame_ods, text='Examinar', bootstyle=PRIMARY, command=self.browse_ods).pack(side=LEFT, padx=5)

        # Delay between emails
        ttk.Label(self, text='Retraso Entre Correos (segundos)').pack(pady=(10, 0))
        self.delay_entry = ttk.Entry(self)
        self.delay_entry.pack(fill=X, padx=20)
        self.delay_entry.insert(0, self.config.get('delay', '10'))

        # Buttons
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text='Guardar Config', bootstyle=SUCCESS, command=self.save_conf).pack(side=LEFT, padx=10)
        self.send_button = ttk.Button(btn_frame, text='Enviar Correos', bootstyle=INFO, command=self.start_sending)
        self.send_button.pack(side=LEFT, padx=10)
        
        # Progress Bar
        ttk.Label(self, text='Progreso:').pack(pady=(10, 0))
        self.progress = ttk.Progressbar(self, bootstyle="success-striped", length=510, mode='determinate')
        self.progress.pack(padx=20, pady=5)
        
        # Console Output
        ttk.Label(self, text='Consola de Salida:').pack(pady=(10, 0))
        self.console = scrolledtext.ScrolledText(self, height=10, width=60)
        self.console.pack(padx=20, pady=5, fill=BOTH, expand=YES)
        self.console.config(state=DISABLED)

    def log_message(self, message):
        """Add message to console with timestamp"""
        self.console.config(state=NORMAL)
        timestamp = time.strftime('%H:%M:%S')
        self.console.insert(END, f"[{timestamp}] {message}\n")
        self.console.see(END)  # Auto-scroll to latest message
        self.console.config(state=DISABLED)
        
        # Process pending events to update UI
        self.update_idletasks()

    def browse_html(self):
        path = filedialog.askopenfilename(filetypes=[('Archivos HTML', '*.html'), ('Todos los Archivos', '*.*')])
        if path:
            self.html_path.set(path)

    def browse_ods(self):
        path = filedialog.askopenfilename(filetypes=[('Archivos ODS', '*.ods'), ('Todos los Archivos', '*.*')])
        if path:
            self.ods_path.set(path)

    def save_conf(self):
        data = {
            'smtp_server': self.smtp_entry.get(),
            'port': self.port_entry.get(),
            'user': self.user_entry.get(),
            'password': self.pw_entry.get(),
            'delay': self.delay_entry.get()
        }
        save_config(data)
        self.log_message('Configuración guardada en config.json')
        messagebox.showinfo('Guardado', 'Configuración guardada en config.json')

    def start_sending(self):
        """Start the email sending process in a separate thread"""
        smtp_server = self.smtp_entry.get()
        port = self.port_entry.get()
        user = self.user_entry.get()
        password = self.pw_entry.get()
        html = self.html_path.get()
        ods = self.ods_path.get()
        delay = self.delay_entry.get()

        # Validate inputs
        if not all([smtp_server, port, user, password, html, ods, delay]):
            messagebox.showwarning('Campos faltantes', 'Por favor complete todos los campos y seleccione los archivos.')
            return

        try:
            port = int(port)
            delay = float(delay)
        except ValueError:
            messagebox.showerror('Entrada inválida', 'El puerto debe ser un número entero y el retraso debe ser un número.')
            return

        # Disable button during sending
        self.send_button.config(state=DISABLED)
        
        # Clear console
        self.console.config(state=NORMAL)
        self.console.delete(1.0, END)
        self.console.config(state=DISABLED)
        
        # Reset progress bar
        self.progress['value'] = 0
        
        # Start email sending in a thread
        thread = threading.Thread(target=self.send_emails, 
                                 args=(smtp_server, port, user, password, html, ods, delay))
        thread.daemon = True
        thread.start()

    def send_emails(self, smtp_server, port, sender_email, password, html_path, ods_path, delay=10):
        """Send emails with progress updates to the UI"""
        # Read HTML content
        try:
            self.log_message('Leyendo plantilla HTML...')
            with open(html_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
        except Exception as e:
            self.log_message(f'Error al leer la plantilla HTML: {e}')
            messagebox.showerror('Error', f'Error al leer la plantilla HTML:\n{e}')
            self.send_button.config(state=NORMAL)
            return

        # Load emails from ODS
        try:
            self.log_message('Leyendo lista de correos del archivo ODS...')
            doc = ezodf.opendoc(ods_path)
            sheet = doc.sheets[0]
            email_list = [row[0].value for row in sheet.rows() if row[0].value and '@' in str(row[0].value)]
            self.log_message(f'Se encontraron {len(email_list)} direcciones de correo')
        except Exception as e:
            self.log_message(f'Error al leer el archivo ODS: {e}')
            messagebox.showerror('Error', f'Error al leer el archivo ODS:\n{e}')
            self.send_button.config(state=NORMAL)
            return

        if not email_list:
            self.log_message('¡No se encontraron direcciones de correo válidas en el archivo!')
            messagebox.showwarning('Advertencia', '¡No se encontraron direcciones de correo válidas en el archivo!')
            self.send_button.config(state=NORMAL)
            return

        # Connect to SMTP server
        try:
            self.log_message(f'Conectando al servidor SMTP {smtp_server}:{port}...')
            server = smtplib.SMTP(smtp_server, port)
            server.ehlo()
            self.log_message('Iniciando cifrado TLS...')
            server.starttls()
            server.ehlo()
            self.log_message('Iniciando sesión...')
            server.login(sender_email, password)
            self.log_message('¡Sesión iniciada con éxito!')
        except Exception as e:
            self.log_message(f'Falló la conexión SMTP: {e}')
            messagebox.showerror('Error', f'Falló la conexión SMTP:\n{e}')
            self.send_button.config(state=NORMAL)
            return

        # Send emails
        sent = 0
        total = len(email_list)
        failed = []

        for i, receiver in enumerate(email_list):
            self.log_message(f'Enviando correo a {receiver}...')
            
            # Update progress bar
            progress_value = int((i / total) * 100)
            self.progress['value'] = progress_value
            
            msg = MIMEMultipart('alternative')
            msg['Subject'] = 'Asunto de su Correo'
            msg['From'] = sender_email
            msg['To'] = receiver
            msg.attach(MIMEText(html_content, 'html'))
            
            try:
                server.sendmail(sender_email, receiver, msg.as_string())
                sent += 1
                self.log_message(f'✓ Correo enviado con éxito a {receiver}')
            except Exception as e:
                failed.append(receiver)
                self.log_message(f'✗ Falló el envío a {receiver}: {e}')
            
            if i < total - 1:  # Don't delay after the last email
                self.log_message(f'Esperando {delay} segundos antes de enviar el siguiente correo...')
                time.sleep(delay)

        # Complete the progress bar
        self.progress['value'] = 100
        
        server.quit()
        self.log_message('Conexión SMTP cerrada')
        
        # Summary
        self.log_message(f'===== RESUMEN =====')
        self.log_message(f'Total de correos: {total}')
        self.log_message(f'Enviados con éxito: {sent}')
        self.log_message(f'Fallidos: {len(failed)}')
        
        if failed:
            self.log_message('Destinatarios fallidos:')
            for email in failed:
                self.log_message(f'  - {email}')
        
        # Re-enable the send button
        self.send_button.config(state=NORMAL)
        
        # Show completion message
        messagebox.showinfo('Completado', f'¡Envío de correos completado!\n\nEnviados con éxito: {sent}/{total}')


if __name__ == '__main__':
    app = EmailSenderApp()
    app.mainloop()