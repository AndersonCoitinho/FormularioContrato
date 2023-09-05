import boto3
import os
from botocore.exceptions import NoCredentialsError

AWS_ACCESS_KEY = os.environ.get('AWS_ACCESS_KEY')
AWS_SECRET_KEY = os.environ.get('AWS_SECRET_KEY')

s3 = boto3.client('s3', aws_access_key_id=AWS_ACCESS_KEY, aws_secret_access_key=AWS_SECRET_KEY)

def upload_to_s3(local_file, bucket_name, s3_file):       
    
    # Fazer upload dos documentos para o S3
    try:
        s3.upload_file(local_file, bucket_name, s3_file)
        print("Upload realizado com sucesso!")
        return True
    except FileNotFoundError:
        print("O arquivo não foi encontrado.")
        return False
    except NoCredentialsError as e:
        print("Erro de credenciais:", e)
        return False
