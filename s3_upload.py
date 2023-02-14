import boto3
import os
from dotenv import load_dotenv
from subprocess import run
from docx2pdf import convert

load_dotenv()

run(['python3', 'main.py', '--local'])

filename = 'cv.pdf'

convert('cv.docx', filename)

if not os.path.exists(filename):
    raise FileNotFoundError("This file does not exist")

with open(filename, 'rb') as f:
    cvfile = f

    session = boto3.Session(
        aws_access_key_id=os.environ.get('AWS_ACCESS_KEY_ID'),
        aws_secret_access_key=os.environ.get('AWS_SECRET_ACCESS_KEY')
    )

    # Creating S3 Resource From the Session.
    s3 = session.resource('s3')

    object = s3.Object(os.environ.get('AWS_STORAGE_BUCKET_NAME'), f'personal-website/{filename}')

    result = object.put(Body=cvfile)

    res = result.get('ResponseMetadata')

    if res.get('HTTPStatusCode') == 200:
        print('File Uploaded Successfully')
    else:
        print('File Not Uploaded')
