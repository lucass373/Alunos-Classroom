from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials

credsClass = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/classroom.courses.readonly', 'https://www.googleapis.com/auth/classroom.rosters.readonly'])

# Criar uma instância do serviço Classroom
classroom_service = build('classroom', 'v1', credentials=credsClass)

# Obter a lista de salas de aula
serviceClass = build('classroom', 'v1', credentials= credsClass)
courses = []
values = []
response = serviceClass.courses().list(pageSize=500).execute()
courses.extend(response.get('courses', []))

#para escrever na planilha
def planilha_escreve(val):
    try:
        # credenciais de autenticação
        creds = service_account.Credentials.from_service_account_file('credentialsPlanilha.json')

        # ID da planilha
        spreadsheet_id = '13zp5oVVi-OSyGACsm2_f-uxRBkOSwz0L84I2nM8ZbCo'

        # criação do cliente
        service = build('sheets', 'v4', credentials=creds)
        arr = []
        # corpo da requisição
        for vals in val:
            arr.append(vals.split('|'))

        body = {
            'values': arr
        }

        # chamada para escrever os valores na planilha
        service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id, range='Alunos que estão no classroom!A1',#Setamos a planilha e a célula em que irá escrever os valores
            valueInputOption='USER_ENTERED', insertDataOption="OVERWRITE", body=body).execute()
        
    except HttpError as http:
        print(f"Error {http}")
#-------------------------FIM DA FUNÇÃO-------------------------------

# CONTAGEM DO NÚMERO DE SALAS NO CLASSROOM
while 'nextPageToken' in response:
    page_token = response['nextPageToken']
    response = serviceClass.courses().list(pageSize=1000, pageToken=page_token).execute()
    courses.extend(response.get('courses', []))

print('Número total de salas de aula:', len(courses))
#-------------------------------------------------

#COMEÇA A COLETA DOS ALUNOS E A PREPARAÇÃO DO ARRAY PARA ESCREVER NO GOOGLE SHEETS
for course in courses:
  
    num_students= 0
    course_id = course['id']
    course_name = course['name']
    
    # Obter a lista de alunos matriculados na sala de aula
    response = serviceClass.courses().students().list(courseId=course_id, pageSize=1000).execute()
    students = response.get('students', [])
    for student in students:
        print(f'{student["profile"]["id"]}|{course_id}')
        values.append(f'{student["profile"]["id"]}|{course_id}') # insere os valores coletados no array


    # Mostrar o número de alunos na sala de aula
    num_students += len(students)
    while 'nextPageToken' in response:
        page_token = response['nextPageToken']
        response = serviceClass.courses().students().list(courseId=course_id, pageSize=1000, pageToken=page_token).execute()
        students = response.get('students', [])
        for student in students:
         print(f'{student["profile"]["id"]}|{course_id}')
         #planilha_escreve(f'{student["profile"]["id"]};{course_id}')
         values.append(f'{student["profile"]["id"]}|{course_id}')
         num_students += len(students)
planilha_escreve(values)# será chamada a função "planilha_escreve", onde o parâmetro serão os valores inseridos no array list "values", nesta função a string será separada pelo "|" e escreverá na planilha setada 
                        # na variável "spreadsheet_id"

