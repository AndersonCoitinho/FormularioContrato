from flask import Flask, render_template, request, send_file
from docx import Document

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import zipfile

from datetime import datetime
import locale


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_docx', methods=['POST'])
def gerar_docx():
    nome = request.form['nome']
    data_str = request.form['data']

    # Converter a data em um objeto datetime
    data = datetime.strptime(data_str, '%Y-%m-%d')
    # Definir a localidade para o idioma desejado (por exemplo, 'pt_BR' para Português do Brasil)
    locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
    # Formatar a data por extenso
    data_extenso = data.strftime('%d de %B de %Y')  # %d: dia, %B: mês por extenso, %Y: ano



    doc1 = Document('./modelos/contratoHonorarios.docx') # Substitua 'modelo.docx' pelo caminho do seu modelo DOCX

    for table in doc1.tables: #percorrendo todas as tabelas
        for row in table.rows: #percorrendo todas as linhas
            for cell in row.cells: #percorrendo todas as cedulas 
                cell_text = cell.text 
                if '{{nome}}' in cell_text: #se encontrar {{NOME}}
                    cell.text = cell_text.replace('{{nome}}', nome) #vai inserir o valor do usuario
               


    doc1_path = (f'Contrato Honorarios_{nome}.docx')
    doc1.save(doc1_path)


    doc2 = Document('./modelos/justicagratuita.docx')
    for table in doc2.tables: #percorrendo todas as tabelas
        for row in table.rows: #percorrendo todas as linhas
            for cell in row.cells: #percorrendo todas as cedulas 
                cell_text = cell.text 
                if '{{nome}}' in cell_text: #se encontrar {{NOME}}
                    cell.text = cell_text.replace('{{nome}}', nome) #vai inserir o valor do usuario
    doc2_path = (f'Justica Gratuita_{nome}.docx')
    doc2.save(doc2_path)



    doc3 = Document('./modelos/procuracao.docx')
    for table in doc3.tables: #percorrendo todas as tabelas
        for row in table.rows: #percorrendo todas as linhas
            for cell in row.cells: #percorrendo todas as cedulas 
                cell_text = cell.text 
                #if '{{nome}}' in cell_text: #se encontrar {{NOME}}
                #    cell.text = cell_text.replace('{{nome}}', nome) #vai inserir o valor do usuario
                if '{{data}}' in cell_text: 
                    cell.text = cell_text.replace('{{data}}', data_extenso)

    doc3_path = (f'Procuracao_{nome}.docx')
    doc3.save(doc3_path)

    #doc2.save(f'Justica Gratuita_{nome}.docx') #salvo o arquivo
    #return send_file(f'Justica Gratuita_{nome}.docx', as_attachment=True)

    # Criar um arquivo ZIP e adicionar os arquivos DOCX
    zip_path = 'arquivos_gerados.zip'
    with zipfile.ZipFile(zip_path, 'w') as zip_file:
        zip_file.write(doc1_path)
        zip_file.write(doc2_path)
        zip_file.write(doc3_path)

    #Enviar o arquivo ZIP para download automático
    return send_file(zip_path, as_attachment=True)



    #doc3 = Document('./modelos/procuracao.docx')

"""
def enviar_email(destinatario_email, arquivo_anexo):
    de_email = 'rocha98918@gmail.com'  # Insira o seu endereço de email
    senha = 'gjwjdybcftipvnbm'  # Insira sua senha de email
    destinatario_email= 'andercoitinho@gmail.com'
    msg = MIMEMultipart()
    msg['From'] = de_email
    msg['To'] = destinatario_email
    msg['Subject'] = 'Arquivo DOCX Gerado e Anexado'

    corpo = f'Olá,\n\nSegue o arquivo DOCX gerado.\n\nAtenciosamente,\nSeu Nome'
    msg.attach(MIMEText(corpo, 'plain'))


    attachment = open(arquivo_anexo, 'rb') #anexando o docx ao email
    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={arquivo_anexo}')
    msg.attach(part)

    
    server = smtplib.SMTP('smtp.gmail.com', 587)  # Insira o servidor SMTP e a porta
    server.starttls() #iniciando uma conexao segura
    server.login(de_email, senha) #fazendo login
    texto = msg.as_string()
    server.sendmail(de_email, destinatario_email, texto)
    server.quit() #finalizando

    """



if __name__ == '__main__':
    app.run(debug=True)
