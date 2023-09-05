import os
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import tkinter as tk
from tkinter import filedialog

def send_email(to_email, subject, body, attachments=None, sender_email=None, password=None, smtp_server=None, port=None):
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if attachments:
        for file_path in attachments:
            part = MIMEBase("application", "octet-stream")
            with open(file_path, "rb") as file:
                part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {os.path.basename(file_path)}"
            )
            msg.attach(part)

    try:
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.sendmail(sender_email, to_email, msg.as_string())
        print(f"Email sent successfully to: {to_email}")
    except Exception as e:
        print(f"Error sending email to: {to_email}")
        print(f"Error details: {e}")

def main():
    # Create a simple GUI using tkinter
    root = tk.Tk()
    root.title("Email Sender")
    
    # Initialize variables
    sender_email = None
    smtp_server = "smtp.gmail.com"  # Replace with your SMTP server address
    port = 587  # Replace with the appropriate port number (e.g., 587 for TLS)
    password = None
    database_file = None
    attachments = []
    
    # Database File Button
    def browse_database():
        nonlocal database_file
        database_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        database_label.config(text=f"Selected Database: {os.path.basename(database_file)}")
    
    database_button = tk.Button(root, text="Origen de los datos", command=browse_database)
    database_button.pack()
    database_label = tk.Label(root, text="Selected Database: None")
    database_label.pack()

    # Sender Email Entry
    sender_email_label = tk.Label(root, text="Email:")
    sender_email_label.pack()
    sender_email_entry = tk.Entry(root)
    sender_email_entry.pack()

    # Password Entry
    password_label = tk.Label(root, text="Contraseña:")
    password_label.pack()
    password_entry = tk.Entry(root, show="*")
    password_entry.pack()

    # Attachments Button
    def browse_files():
        nonlocal attachments
        file_paths = filedialog.askopenfilenames()
        attachments.extend(file_paths)
        update_attachments_listbox()
    
    attachments_button = tk.Button(root, text="Ficheros Adjuntos", command=browse_files)
    attachments_button.pack()

    # Attachments Listbox
    attachments_listbox = tk.Listbox(root, selectmode=tk.SINGLE)
    attachments_listbox.pack()
    attachments_listbox.config(width=50)

    def update_attachments_listbox():
        attachments_listbox.delete(0, tk.END)
        for file in attachments:
            attachments_listbox.insert(tk.END, os.path.basename(file))

    # Delete Attachments Button
    def delete_attachments():
        selected_attachment_index = attachments_listbox.curselection()
        if selected_attachment_index:
            index = int(selected_attachment_index[0])
            del attachments[index]
            update_attachments_listbox()

    delete_attachments_button = tk.Button(root, text="Eliminar Adjunto", command=delete_attachments)
    delete_attachments_button.pack()

    # Send Button
    def send_emails():
        nonlocal sender_email, smtp_server, port, password, database_file  # Use the variables from the outer scope
        sender_email = sender_email_entry.get()
        password = password_entry.get()
        data = pd.read_excel(database_file)
        for index, row in data.iterrows():
            name = row["Nombre"]
            last_name = row["Apellido"]
            street = row["Calle"]
            situation = row["Situación"]
            mail_address = row["Dirección"]
            email_body = email_template.format(Nombre=name, Apellido=last_name, Calle=street, Situación=situation)
            send_email(mail_address, "Admission Status", email_body, attachments=attachments, sender_email=sender_email, password=password, smtp_server=smtp_server, port=port)
    
        root.destroy()  # Close the tkinter window after emails are sent

    send_button = tk.Button(root, text="Enviar Correos", command=send_emails)
    send_button.pack()

    root.mainloop()

if __name__ == "__main__":
    email_template = """
    Estimado/a {Nombre} {Apellido},

    Nos complace informarle que su solicitud ha sido {Situación}.

    A continuación, los detalles que tenemos registrados:
    Calle: {Calle}

    Si tiene alguna pregunta, no dude en ponerse en contacto con nosotros.

    Atentamente,
    Su Organización
    """
    main()