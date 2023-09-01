from __future__ import print_function
from flask import Flask, render_template, request, redirect, make_response, session
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
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from utils.date_utils import format_data_extenso
from utils.upload_s3 import upload_to_s3

app = Flask(__name__)

# Configurações de autenticação
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


@app.route('/')
def index():
    return render_template('index.html')


AWS_ACCESS_KEY = "AKIAYGU4WWO6VQB4AQUO"
AWS_SECRET_KEY = "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"


@app.route('/generate_docx', methods=['POST'])
def gerar_docx():
    try:
        nome = request.form['nome'].upper()
        nacionalidade = request.form['nacionalidade'].upper()
        estadoCivil = request.form['estadoCivil'].upper()
        profissao = request.form['profissao'].upper()
        fone = request.form['fone'].upper()
        fone_recado = request.form['fone_recado'].upper()
        cpf = request.form['cpf'].upper()
        rg = request.form['rg'].upper()
        endereco = request.form['endereco'].upper()
        bairro = request.form['bairro'].upper()
        cep = request.form['cep'].upper()
        cidade = request.form['cidade'].upper()
        estado = request.form['estado'].upper()
        cep = request.form['cep'].upper()
        data_str = request.form['data']

        # Chamando a função format_data_extenso para obter a data por extenso
        data_extenso = format_data_extenso(data_str)
        

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
                        if fone_recado:
                            cell.text = cell_text.replace('{{fone}}', fone + f' ou {fone_recado}')
                        else:
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
        
        for paragraph in doc1.paragraphs: #percorre os paragratos
            paragraph_text = paragraph.text
            if '{{data}}' in paragraph_text: #encontrando {{data}} substitui
                paragraph.text = paragraph_text.replace('{{data}}', data_extenso)

        doc1_path = os.path.join('modelos', f'CONTRATO HONORÁRIO - {nome}.docx')
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
                        if fone_recado:
                            cell.text = cell_text.replace('{{fone}}', fone + f' ou {fone_recado}')
                        else:
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
        
        for paragraph in doc2.paragraphs: #percorre os paragratos
            paragraph_text = paragraph.text
            if '{{data}}' in paragraph_text: #encontrando {{data}} substitui
                paragraph.text = paragraph_text.replace('{{data}}', data_extenso)

        doc2_path = os.path.join('modelos', f'JUSTIÇA GRATUITA - {nome}.docx')
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
                        if fone_recado:
                            cell.text = cell_text.replace('{{fone}}', fone + f' ou {fone_recado}')
                        else:
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
        
        for paragraph in doc3.paragraphs: #percorre os paragratos
            paragraph_text = paragraph.text
            if '{{data}}' in paragraph_text: #encontrando {{data}} substitui
                paragraph.text = paragraph_text.replace('{{data}}', data_extenso)

        doc3_path = os.path.join('modelos', f'PROCURAÇÃO - {nome}.docx')
        doc3.save(doc3_path)
       
        ### DOC4 = Capa Processo ###
        doc4 = Document('./modelos/capaProcesso.docx')
        for table in doc4.tables: #percorrendo todas as tabelas
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
                        if fone_recado:
                            cell.text = cell_text.replace('{{fone}}', fone + f' ou {fone_recado}')
                        else:
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

        doc4_path = os.path.join('modelos', f'CAPA DO PROCESSO - {nome}.docx')
        doc4.save(doc4_path)


        ### DOC5 = Minuta Auxilio Acidente Federal ###
        doc5 = Document('./modelos/minutaAuxilioAcidenteFederal.docx')
           
        for paragraph in doc5.paragraphs:
            paragraph_text = paragraph.text
            if '{{nome}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{nome}}', nome)
            if '{{nacionalidade}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{nacionalidade}}', nacionalidade) 
            if '{{estadoCivil}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{estadoCivil}}', estadoCivil)
            if '{{profissao}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{profissao}}', profissao)
            if '{{cpf}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{cpf}}', cpf)
            if '{{rg}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{rg}}', rg)
            if '{{endereco}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{endereco}}', endereco)
            if '{{bairro}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{bairro}}', bairro)
            if '{{cidade}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{cidade}}', cidade)
            if '{{estado}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{estado}}', estado)
            if '{{cep}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{cep}}', cep)
            if '{{data}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{data}}', data_extenso)
   
            # Atribuir o texto modificado de volta ao parágrafo
            paragraph.text = paragraph_text

        doc5_path = os.path.join('modelos', f'MINUTA AUXILIO ACIDENTE FEDERAL - {nome}.docx')
        doc5.save(doc5_path)


        ### DOC6 = Requerimento Adm Auxilio Acidente ###
        doc6 = Document('./modelos/requerimentoAdmAuxilioAcidente.docx')

        for paragraph in doc6.paragraphs:
            paragraph_text = paragraph.text
            if '{{nome}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{nome}}', nome)
            if '{{nacionalidade}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{nacionalidade}}', nacionalidade)
            if '{{estadoCivil}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{estadoCivil}}', estadoCivil)
            if '{{profissao}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{profissao}}', profissao)
            if '{{cpf}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{cpf}}', cpf)
            if '{{rg}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{rg}}', rg)
            if '{{endereco}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{endereco}}', endereco)
            if '{{bairro}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{bairro}}', bairro)
            if '{{cidade}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{cidade}}', cidade)
            if '{{estado}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{estado}}', estado)
            if '{{cep}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{cep}}', cep)
            if '{{data}}' in paragraph_text:
                paragraph_text = paragraph_text.replace('{{data}}', data_extenso)

            # Atribuir o texto modificado de volta ao parágrafo
            paragraph.text = paragraph_text

        doc6_path = os.path.join('modelos', f'REQUERIMENTO ADMINISTRATIVO AUXILIO ACIDENTE - {nome}.docx')
        doc6.save(doc6_path)

        """
        # Fazer upload dos documentos para o S3
        if upload_to_s3(doc1_path, 'cadastroadv', f'datas/CONTRATO HONORÁRIO - {nome}.docx') and \
           upload_to_s3(doc2_path, 'cadastroadv', f'datas/JUSTIÇA GRATUITA - {nome}.docx') and \
           upload_to_s3(doc3_path, 'cadastroadv', f'datas/PROCURAÇÃO - {nome}.docx') and \
           upload_to_s3(doc4_path, 'cadastroadv', f'datas/CAPA DO PROCESSO - {nome}.docx') and \
           upload_to_s3(doc5_path, 'cadastroadv', f'datas/MINUTA AUXILIO ACIDENTE FEDERAL - {nome}.docx') and \
           upload_to_s3(doc6_path, 'cadastroadv', f'datas/REQUERIMENTO ADMINISTRATIVO AUXILIO ACIDENTE - {nome}.docx'):
           return redirect(url_for('download_files', nome=nome))
          #return "Documentos gerados e enviados com sucesso!"
        else:
            return "Erro ao gerar e/ou enviar os documentos."
        
            try:
                    update_result = update_spreadsheet_values(
                        service,  # Passe a instância do serviço do Google Sheets aqui
                        "1nBhothHfyCnMgj7egotQapYXlNqXRPoseur1idUY9eE",  # ID da planilha
                        "page!A15",  # Intervalo onde deseja atualizar os valores
                        valores_adicionar  # Valores a serem adicionados
                    )
                    if update_result:
                        return redirect(url_for('download_files', nome=nome))
                    else:
                        return "Erro ao gerar e/ou enviar os documentos."
            except Exception as e:
                return f"Erro inesperado na atualização da planilha: {str(e)}"
            else:
                return "Erro ao gerar e/ou enviar os documentos."
        
        creds = None #credencial vazio

        if os.path.exists('token.json'): #se existe o token.json
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token: # se ja foi atualizado manual ele vai permitir sempre
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            service = build('sheets', 'v4', credentials=creds) #conectando no google sheet

                # Call the Sheets API
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId="1nBhothHfyCnMgj7egotQapYXlNqXRPoseur1idUY9eE",
                                            range="page!A1:B3").execute()
                
                #utilizado para pegar valores
                #values = result.get('values', []) #pega os valores
                #print(values)

                #utilizar para inserir valores
            valores_adicionar = [
            [nome, nacionalidade, estadoCivil, profissao],
            ]

            return sheet.values().update(spreadsheetId="1nBhothHfyCnMgj7egotQapYXlNqXRPoseur1idUY9eE",
                                            range="page!a15", valueInputOption="USER_ENTERED", body={"values": valores_adicionar}).execute()
        
        except HttpError as err:
            return print(err) 
        """
        """
        # Fazer upload dos documentos para o S3
        s3_upload_success = upload_to_s3(doc1_path, 'cadastroadv', f'datas/CONTRATO HONORÁRIO - {nome}.docx') and \
                            upload_to_s3(doc2_path, 'cadastroadv', f'datas/JUSTIÇA GRATUITA - {nome}.docx') and \
                            upload_to_s3(doc3_path, 'cadastroadv', f'datas/PROCURAÇÃO - {nome}.docx') and \
                            upload_to_s3(doc4_path, 'cadastroadv', f'datas/CAPA DO PROCESSO - {nome}.docx') and \
                            upload_to_s3(doc5_path, 'cadastroadv', f'datas/MINUTA AUXILIO ACIDENTE FEDERAL - {nome}.docx') and \
                            upload_to_s3(doc6_path, 'cadastroadv', f'datas/REQUERIMENTO ADMINISTRATIVO AUXILIO ACIDENTE - {nome}.docx')

        if s3_upload_success:
            
            creds = None  # Inicializa a variável 'creds' como None
            
            if os.path.exists('token.json'):
                creds = Credentials.from_authorized_user_file('token.json', SCOPES)
            
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    # Atualize a URL de redirecionamento aqui
                    flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES, 
                                                                     redirect_uri='https://conversordocx-b549802ec5d8.herokuapp.com')
                    creds = flow.run_local_server(port=0)
 
                    
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())
            
            try:
                service = build('sheets', 'v4', credentials=creds)

                # Call the Sheets API
                sheet = service.spreadsheets()

                #Encontre a última linha preenchida
                result = sheet.values().get(
                    spreadsheetId="1nBhothHfyCnMgj7egotQapYXlNqXRPoseur1idUY9eE",
                    range="page!D:D",  # Coluna onde você quer verificar a última linha preenchida
                ).execute()
                

                values = result.get('values', [])
                last_filled_row = len(values) + 1  # Próxima linha vazia

                valores_adicionar = [
                    [
                    nome, 
                    estadoCivil, 
                    profissao, 
                    fone, 
                    fone_recado, 
                    cpf, 
                    rg, 
                    endereco, 
                    bairro, 
                    cidade, 
                    estado, 
                    cep
                    ],
                ]

                update_result = sheet.values().update(
                    spreadsheetId="1nBhothHfyCnMgj7egotQapYXlNqXRPoseur1idUY9eE",
                    range=f"page!D{last_filled_row}",
                    valueInputOption="USER_ENTERED",
                    body={"values": valores_adicionar}
                ).execute()

                if update_result:
                    return redirect(url_for('download_files', nome=nome))
                else:
                    return "Erro ao gerar e/ou enviar os documentos."
            except Exception as e:
                return f"Erro inesperado na atualização da planilha: {str(e)}"

        else:
            return f"Erro ao gerar e/ou enviar os documentos."
        """

        # Fazer upload dos documentos para o S3
        if upload_to_s3(doc1_path, 'cadastroadv', f'datas/CONTRATO HONORÁRIO - {nome}.docx') and \
           upload_to_s3(doc2_path, 'cadastroadv', f'datas/JUSTIÇA GRATUITA - {nome}.docx') and \
           upload_to_s3(doc3_path, 'cadastroadv', f'datas/PROCURAÇÃO - {nome}.docx') and \
           upload_to_s3(doc4_path, 'cadastroadv', f'datas/CAPA DO PROCESSO - {nome}.docx') and \
           upload_to_s3(doc5_path, 'cadastroadv', f'datas/MINUTA AUXILIO ACIDENTE FEDERAL - {nome}.docx') and \
           upload_to_s3(doc6_path, 'cadastroadv', f'datas/REQUERIMENTO ADMINISTRATIVO AUXILIO ACIDENTE - {nome}.docx'):
           return redirect(url_for('download_files', nome=nome))
           #return "Documentos gerados e enviados com sucesso!"
        else:
            return f"Erro ao gerar e/ou enviar os documentos."

            
    except KeyError as e:
        return f"Erro: O campo '{e.args[0]}' não foi encontrado nos dados enviados."
    except Exception as e:
        return f"Erro inesperado: {str(e)}"



@app.route('/downloads/<nome>')
def download_files(nome):
    s3 = boto3.client('s3', aws_access_key_id=AWS_ACCESS_KEY, aws_secret_access_key=AWS_SECRET_KEY)

    filenames = [f'CONTRATO HONORÁRIO - {nome}.docx', 
                    f'JUSTIÇA GRATUITA - {nome}.docx', 
                    f'PROCURAÇÃO - {nome}.docx', 
                    f'CAPA DO PROCESSO - {nome}.docx', 
                    f'MINUTA AUXILIO ACIDENTE FEDERAL - {nome}.docx',
                    f'REQUERIMENTO ADMINISTRATIVO AUXILIO ACIDENTE - {nome}.docx'
                    ]
    
    download_links = []
    
    try:
        for filename in filenames:
            url = s3.generate_presigned_url('get_object',
                                           Params={'Bucket': 'cadastroadv', 'Key': f'datas/{filename}'},
                                           ExpiresIn=3600)
            download_links.append({'filename': filename, 'download_link': url})
            
        return render_template('download.html', download_links=download_links)
    except NoCredentialsError:
        return "Credenciais do AWS não foram configuradas."



if __name__ == '__main__':
    #app.run(debug=True)
    app.run(host='0.0.0.0', port=os.environ.get('PORT', 5000))