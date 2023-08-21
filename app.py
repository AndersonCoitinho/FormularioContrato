from flask import Flask, render_template, request, send_file, make_response
from docx import Document
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import zipfile

from datetime import datetime
import locale

from urllib.parse import quote
from flask import url_for

import boto3
from botocore.exceptions import NoCredentialsError


app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate_docx', methods=['POST'])
def gerar_docx():
    nome = request.form['nome']
    nacionalidade = request.form['nacionalidade']
    estadoCivil = request.form['estadoCivil']
    profissao = request.form['profissao']
    fone = request.form['fone']
    cpf = request.form['cpf']
    rg = request.form['rg']
    endereco = request.form['endereco']
    bairro = request.form['bairro']
    cep = request.form['cep']
    cidade = request.form['cidade']
    estado = request.form['estado']
    cep = request.form['cep']
    data_str = request.form['data']


    ### DATA ###
    # Converter a data em um objeto datetime
    #data = datetime.strptime(data_str, '%Y-%m-%d')
    # Definir a localidade para o idioma desejado (por exemplo, 'pt_BR' para Português do Brasil)
    #locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
    # Formatar a data por extenso
    #data_extenso = data.strftime('%d de %B de %Y')  # %d: dia, %B: mês por extenso, %Y: ano


    ### DOC1 = CONTRATO HONORARIOS ###
    doc1 = Document('./modelos/contratoHonorarios.docx') # Substitua 'modelo.docx' pelo caminho do seu modelo DOCX

    for table in doc1.tables: #percorrendo todas as tabelas
        for row in table.rows: #percorrendo todas as linhas
            for cell in row.cells: #percorrendo todas as cedulas 
                cell_text = cell.text 
                if '{{nome}}' in cell_text: #se encontrar {{NOME}}
                    cell.text = cell_text.replace('{{nome}}', nome) #vai inserir o valor do usuario
                if '{{nacionalidade}}' in cell_text:
                    cell.text = cell_text.replace('{{nacionalidade}}', nacionalidade)
                if '{{estadoCivil}}' in cell_text:
                    cell.text = cell_text.replace('{{estadoCivil}}', estadoCivil)
                if '{{profissao}}' in cell_text:
                    cell.text = cell_text.replace('{{profissao}}', profissao)
                if '{{fone}}' in cell_text:
                    cell.text = cell_text.replace('{{fone}}', fone)
                if '{{cpf}}' in cell_text:
                    cell.text = cell_text.replace('{{cpf}}', cpf)
                if '{{rg}}' in cell_text:
                    cell.text = cell_text.replace('{{rg}}', rg)
                if '{{endereco}}' in cell_text:
                    cell.text = cell_text.replace('{{endereco}}', endereco)
                if '{{bairro}}' in cell_text:
                    cell.text = cell_text.replace('{{bairro}}', bairro)
                if '{{cep}}' in cell_text:
                    cell.text = cell_text.replace('{{cep}}', cep)
                if '{{cidade}}' in cell_text:
                    cell.text = cell_text.replace('{{cidade}}', cidade)
                if '{{estado}}' in cell_text:
                    cell.text = cell_text.replace('{{estado}}', estado)
    
    #for paragraph in doc1.paragraphs: #percorre os paragratos
    #    paragraph_text = paragraph.text
    #    if '{{data}}' in paragraph_text: #encontrando {{data}} substitui
    #        paragraph.text = paragraph_text.replace('{{data}}', data_extenso)

    doc1_path = os.path.join('modelos', f'Contrato_Honorarios_{nome}.docx')
    doc1.save(doc1_path)

    ### DOC2 = JUSTIÇA GRATUIDA ###
    doc2 = Document('./modelos/justicagratuita.docx')
    for table in doc2.tables: #percorrendo todas as tabelas
        for row in table.rows: #percorrendo todas as linhas
            for cell in row.cells: #percorrendo todas as cedulas 
                cell_text = cell.text 
                if '{{nome}}' in cell_text: #se encontrar {{NOME}}
                    cell.text = cell_text.replace('{{nome}}', nome) #vai inserir o valor do usuario
                if '{{nacionalidade}}' in cell_text:
                    cell.text = cell_text.replace('{{nacionalidade}}', nacionalidade)
                if '{{estadoCivil}}' in cell_text:
                    cell.text = cell_text.replace('{{estadoCivil}}', estadoCivil)
                if '{{profissao}}' in cell_text:
                    cell.text = cell_text.replace('{{profissao}}', profissao)
                if '{{fone}}' in cell_text:
                    cell.text = cell_text.replace('{{fone}}', fone)
                if '{{cpf}}' in cell_text:
                    cell.text = cell_text.replace('{{cpf}}', cpf)
                if '{{rg}}' in cell_text:
                    cell.text = cell_text.replace('{{rg}}', rg)
                if '{{endereco}}' in cell_text:
                    cell.text = cell_text.replace('{{endereco}}', endereco)
                if '{{bairro}}' in cell_text:
                    cell.text = cell_text.replace('{{bairro}}', bairro)
                if '{{cep}}' in cell_text:
                    cell.text = cell_text.replace('{{cep}}', cep)
                if '{{cidade}}' in cell_text:
                    cell.text = cell_text.replace('{{cidade}}', cidade)
                if '{{estado}}' in cell_text:
                    cell.text = cell_text.replace('{{estado}}', estado)
    
    #for paragraph in doc2.paragraphs: #percorre os paragratos
    #    paragraph_text = paragraph.text
    #    if '{{data}}' in paragraph_text: #encontrando {{data}} substitui
    #        paragraph.text = paragraph_text.replace('{{data}}', data_extenso)

    doc2_path = os.path.join('modelos', f'Justica_Gratuita_{nome}.docx')
    doc2.save(doc2_path)

    ### DOC3 = PROCURAÇÃO ###
    doc3 = Document('./modelos/procuracao.docx')
    for table in doc3.tables: #percorrendo todas as tabelas
        for row in table.rows: #percorrendo todas as linhas
            for cell in row.cells: #percorrendo todas as cedulas 
                cell_text = cell.text 
                if '{{nome}}' in cell_text: #se encontrar {{NOME}}
                    cell.text = cell_text.replace('{{nome}}', nome) #vai inserir o valor do usuario
                if '{{nacionalidade}}' in cell_text:
                    cell.text = cell_text.replace('{{nacionalidade}}', nacionalidade)
                if '{{estadoCivil}}' in cell_text:
                    cell.text = cell_text.replace('{{estadoCivil}}', estadoCivil)
                if '{{profissao}}' in cell_text:
                    cell.text = cell_text.replace('{{profissao}}', profissao)
                if '{{fone}}' in cell_text:
                    cell.text = cell_text.replace('{{fone}}', fone)
                if '{{cpf}}' in cell_text:
                    cell.text = cell_text.replace('{{cpf}}', cpf)
                if '{{rg}}' in cell_text:
                    cell.text = cell_text.replace('{{rg}}', rg)
                if '{{endereco}}' in cell_text:
                    cell.text = cell_text.replace('{{endereco}}', endereco)
                if '{{bairro}}' in cell_text:
                    cell.text = cell_text.replace('{{bairro}}', bairro)
                if '{{cep}}' in cell_text:
                    cell.text = cell_text.replace('{{cep}}', cep)
                if '{{cidade}}' in cell_text:
                    cell.text = cell_text.replace('{{cidade}}', cidade)
                if '{{estado}}' in cell_text:
                    cell.text = cell_text.replace('{{estado}}', estado)
    
    #for paragraph in doc3.paragraphs: #percorre os paragratos
    #    paragraph_text = paragraph.text
    #    if '{{data}}' in paragraph_text: #encontrando {{data}} substitui
    #        paragraph.text = paragraph_text.replace('{{data}}', data_extenso)

    doc3_path = os.path.join('modelos', f'Procuracao_{nome}.docx')
    doc3.save(doc3_path)

    #return render_template(
    #'download.html',
    #doc1_path=doc1_path,
    #doc2_path=doc2_path,
    #doc3_path=doc3_path)

       # Configurar as credenciais do S3
    AWS_ACCESS_KEY = os.environ['AWS_ACCESS_KEY']
    AWS_SECRET_KEY = os.environ['AWS_SECRET_KEY']
    s3 = boto3.client('s3', aws_access_key_id=AWS_ACCESS_KEY, aws_secret_access_key=AWS_SECRET_KEY)

    # Fazer upload dos documentos para o S3
    def upload_to_s3(local_file, bucket_name, s3_file):
        try:
            s3.upload_file(local_file, bucket_name, s3_file)
            print("Upload realizado com sucesso!")
            return True
        except FileNotFoundError:
            print("O arquivo não foi encontrado.")
            return False
        except NoCredentialsError:
            print("Credenciais do AWS não foram configuradas.")
            return False

    # Fazer upload dos documentos para o S3
    upload_to_s3(doc1_path, 'cadastroadv', f'datas/Contrato_Honorarios_{nome}.docx')
    upload_to_s3(doc2_path, 'cadastroadv', f'datas/Justica_Gratuita_{nome}.docx')
    upload_to_s3(doc3_path, 'cadastroadv', f'datas/Procuracao_{nome}.docx')

    return render_template(
        'download.html',
        doc1_path=doc1_path,
        doc2_path=doc2_path,
        doc3_path=doc3_path
    )

if __name__ == '__main__':
    #app.run(debug=True)
    app.run(host='0.0.0.0', port=os.environ.get('PORT', 5000))