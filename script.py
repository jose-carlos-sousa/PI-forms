
import requests
import pandas as pd
from datetime import datetime
import io
import subprocess
from dotenv import load_dotenv
import os
import openpyxl
import urllib.parse

load_dotenv()

LOGIN_URL = os.getenv("LOGIN_URL")
LOGIN_PAYLOAD = {
    "pv_login": os.getenv("LOGIN_USERNAME"),
    "pv_password": os.getenv("LOGIN_PASSWORD")
}

DATA_URLS = {
    "individual": "https://docs.google.com/spreadsheets/d/1ReF6mYhuEzwKrWYYOxXkhwaELv9F-HL_F_Xc4ab7tsI/gviz/tq?tqx=out:csv&gid=965489064",
    "groups": "https://docs.google.com/spreadsheets/d/1ReF6mYhuEzwKrWYYOxXkhwaELv9F-HL_F_Xc4ab7tsI/gviz/tq?tqx=out:csv&gid=0",
    "group_students": "https://docs.google.com/spreadsheets/d/1ANIqJFFmZPyhkEMd4ncuCFnvrAj_sDd3r4MH74Musck/gviz/tq?tqx=out:csv&gid=578405051"
}

session = requests.Session()
response = session.post(LOGIN_URL, data=LOGIN_PAYLOAD)
response.raise_for_status()  
current_week = datetime.now().isocalendar()[1] - 6

def fetch_csv(url):
    response = session.get(url)
    response.raise_for_status()
    return pd.read_csv(io.StringIO(response.text))

def fetch_student_map():
    response = session.get("https://sigarra.up.pt/feup/pt/mob_ucurr_geral.uc_inscritos?pv_ocorrencia_id=541895")
    response.raise_for_status()
    data = response.json()
    student_map = {student['nome']: student['codigo'] for student in data}
    return student_map

def get_individual_students_who_didnt_answer():
    df = fetch_csv(DATA_URLS["individual"])
    filtered_df = df.filter(regex=f'^(Estudante|W{current_week})')
    filtered_df = filtered_df.loc[filtered_df[f'W{current_week} '] != "ok"]
    return filtered_df['Estudante '].tolist()

def get_groups_who_didnt_answer():
    df = fetch_csv(DATA_URLS["groups"])
    filtered_df = df.filter(regex=f'^(Código|W{current_week})')
    filtered_df = filtered_df.loc[filtered_df[f'W{current_week} '] != "ok"]
    return [codigo.split()[0] for codigo in filtered_df['Código '].tolist()]

def get_group_students_who_didnt_answer(groups_who_didnt_answer):
    df = fetch_csv(DATA_URLS["group_students"])
    filtered_df = df[['Código', 'Equipa']]
    filtered_df = filtered_df[filtered_df['Código'].isin(groups_who_didnt_answer)]
    
    students_who_didnt_answer = []
    for equipa in filtered_df['Equipa']:
        students_who_didnt_answer.extend([name.split('(')[0].strip() for name in equipa.split("\n")])
    
    return students_who_didnt_answer

def get_emails(students_who_didnt_answer, student_map):
    emails = []
    for student in students_who_didnt_answer:
        if student in student_map:
            up_code = student_map[student]
            email = f"up{up_code}@up.pt"
            emails.append(email)
    return emails
def parse_excel(file_path):
    df = pd.read_excel(file_path, sheet_name=None)
    parsed_data = {}
    for sheet_name, sheet_data in df.items():
        parsed_data[sheet_name] = sheet_data.to_dict(orient='records')
    return parsed_data

excel_data = parse_excel('excel.xlsx')
allowed_emails = [entry['Email'] for entry in excel_data['Sheet1'] if entry.get('Gostarias de receber um mail sexta de tarde caso não tenhas preenchido o forms?\n') == 'sim']
print(allowed_emails)
individual_students_who_didnt_answer = get_individual_students_who_didnt_answer()
groups_who_didnt_answer = get_groups_who_didnt_answer()
group_students_who_didnt_answer = get_group_students_who_didnt_answer(groups_who_didnt_answer)
students_who_didnt_answer = individual_students_who_didnt_answer + group_students_who_didnt_answer
student_up_map = fetch_student_map()
emails = get_emails(students_who_didnt_answer, student_up_map)
filtered_emails = [email for email in emails if email in allowed_emails]
email_list = ", ".join(filtered_emails)

subject = "Não te esqueças de pi"
body = "Olá, só para avisar que ainda não preencheste o forms semanal de PI. Podes preencher neste link: https://docs.google.com/forms/d/e/1FAIpQLSc9vBdB9inxrSmprhU-DZVbBvI_A6FLqg8I5ucvU2TUtCwt4w/viewform"

encoded_subject = urllib.parse.quote(subject)
encoded_body = urllib.parse.quote(body)

mailto_link = f"mailto:{email_list}?subject={encoded_subject}&body={encoded_body}"

subprocess.run(["xdg-open", mailto_link])

print(emails)
