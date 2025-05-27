from flask import Flask, render_template, request, redirect
import pandas as pd
from openpyxl import load_workbook
import os
import smtplib
from email.message import EmailMessage

app = Flask(__name__)
EXCEL_FILE = 'data.xlsx'
EMAIL_SENDER = 'ing.patriciovaldes@gmail.com'
EMAIL_PASSWORD = 'zxqx fhlo puut weoi'
EMAIL_RECEIVER = 'ing.patriciovaldes@gmail.com'

def create_excel_if_not_exists():
    if not os.path.exists(EXCEL_FILE):
        df = pd.DataFrame(columns=["Agricultor", "Labor", "Fecha", "Tipo de Cultivo"])
        df.to_excel(EXCEL_FILE, index=False)

def add_data(agricultor, labor, fecha, cultivo):
    df = pd.read_excel(EXCEL_FILE)
    new_row = {"Agricultor": agricultor, "Labor": labor, "Fecha": fecha, "Tipo de Cultivo": cultivo}
    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
    df.to_excel(EXCEL_FILE, index=False)

def send_email_with_excel():
    msg = EmailMessage()
    msg['Subject'] = 'Actualización de datos agrícolas'
    msg['From'] = EMAIL_SENDER
    msg['To'] = EMAIL_RECEIVER
    msg.set_content('Se ha actualizado el Excel con nueva información.')

    with open(EXCEL_FILE, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=EXCEL_FILE)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
        smtp.send_message(msg)

@app.route('/', methods=['GET', 'POST'])
def form():
    if request.method == 'POST':
        agricultor = request.form['agricultor']
        labor = request.form['labor']
        fecha = request.form['fecha']
        cultivo = request.form['cultivo']

        create_excel_if_not_exists()
        add_data(agricultor, labor, fecha, cultivo)
        send_email_with_excel()

        return redirect('/')
    return render_template('form.html')

if __name__ == '__main__':
    create_excel_if_not_exists()
    app.run(debug=True)
