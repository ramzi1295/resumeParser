import re
import sys
import importlib
importlib.reload(sys)
from tika import parser
import spacy
from spacy.matcher import Matcher
import docx2txt
import unidecode
from pyresparser import ResumeParser
import os
import comtypes.client
import psycopg2



nlp = spacy.load('fr_core_news_md')
matcher = Matcher(nlp.vocab)


def extract_text_from_doc(doc_path):
    temp = docx2txt.process(doc_path)
    return temp

def extract_text_from_pdf(pdf_path):
    raw = parser.from_file(pdf_path)
    txt = raw['content']
    return txt

def extractEmail(mytext):
    pattern = re.compile(r'([\w0-9-._]+@[\w0-9-.]+[\w0-9]{2,3})')
    matches = pattern.finditer(mytext)
    for match in matches:
      # print(type(match))
      # if(numberOfDigits(match.group(0)) >=8):
          return (format(match.group(0)))
def extractPhone(mytext):
    pattern = re.compile(r'[+0-900][0-9 (]*[0-9 ]*[0-9 )]*[0-9][0-9 ][0-9][0-9 ][0-9][0-9 ]{4,20}')
    matches = pattern.finditer(mytext)
    for match in matches:
      # print(type(match))
      # if(numberOfDigits(match.group(0)) >=8):
          return (format(match.group(0)))

def docxToPdf(pathIn, pathOut):
        wdFormatPDF = 17
        in_file = os.path.abspath(pathIn)
        out_file = os.path.abspath(pathOut)

        word = comtypes.client.CreateObject('Word.Application')
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close()
        word.Quit()


def extractExperience(path):
    data = ResumeParser(path).get_extracted_data()
    try:
        return data['experience']
    except:
        try:
            return data['experiences']
        except:
            try:
                return data['parcours']
            except:
                try:
                    return data['EXPÉRIENCE']
                except:
                    try:
                        return data['career']
                    except:
                        try:
                            return data['expertise']
                        except:
                            try:
                                return data['job']
                            except:
                                try:
                                    return data['work']
                                except:
                                    try:
                                        return data['internships']
                                    except:
                                        try:
                                            return data['internship']
                                        except:
                                            return None


def extract_name(text):
    nlp_text = nlp(text)
    # First name and Last name are always Proper Nouns
    pattern = [{'POS': 'PROPN'}, {'POS': 'PROPN'}]

    matcher.add('NAME', None, pattern)

    matches = matcher(nlp_text)

    lowerText = text.lower()
    lowerText = unidecode.unidecode(lowerText)
    if 'nom' and 'prenom' in lowerText:
        if lowerText.find('nom') < lowerText.find('prenom'):
            nom = text[
                  lowerText.find(':', lowerText.find('nom') + 3) + 1:lowerText.find('\n', lowerText.find('nom') + 6)]
            prenom = text[lowerText.find(':', lowerText.find('prenom') + 6) + 1:lowerText.find('\n', lowerText.find(':',
                                                                                                                    lowerText.find(
                                                                                                                        'prenom') + 6))]
            return nom + prenom
        elif lowerText.find('nom') > lowerText.find('prenom'):
            prenom = text[lowerText.find(':', lowerText.find('prenom') + 6) + 1:lowerText.find('\n', lowerText.find(
                'prenom') + 8)]
            nom = text[
                  lowerText.find(':', lowerText.find('nom', lowerText.find('prenom') + 10)) + 1:lowerText.find('\n',
                                                                                                               lowerText.find(
                                                                                                                   ':',
                                                                                                                   lowerText.find(
                                                                                                                       'nom',
                                                                                                                       lowerText.find(
                                                                                                                           'prenom') + 10)))]
            return nom + prenom
    elif 'last name' and 'first name' in lowerText:
        nom = text[lowerText.find(':', lowerText.find('last name')) + 1:lowerText.find('\n', lowerText.find(':',
                                                                                                            lowerText.find(
                                                                                                                'last name')))]
        prenom = text[lowerText.find(':', lowerText.find('first name')) + 1:lowerText.find('\n', lowerText.find(':',
                                                                                                                lowerText.find(
                                                                                                                    'first name')))]
        return nom + prenom
    elif 'full name' in lowerText:
        return lowerText[lowerText.find(':', lowerText.find('full name')) + 1:lowerText.find('\n', lowerText.find(':',
                                                                                                                  lowerText.find(
                                                                                                                      'full name')))]
    else:
        for match_id, start, end in matches:
            span = nlp_text[start:end]
            return span.text

def insert(name , phone , email, experience, skills):
    postgres_insert_query = """ INSERT INTO candidate (	"fullName", "phone", "email", "experience", "skills") VALUES (%s,%s,%s,%s,%s)"""
    record_to_insert = (name, phone, email, experience, skills)
    cursor.execute(postgres_insert_query, record_to_insert)
    connection.commit()

try:
    connection = psycopg2.connect(user="postgres",
                                  password="Z",
                                  host="127.0.0.1",
                                  port="5432",
                                  database="Pentabell1")
    cursor = connection.cursor()
    print('Connected!')
except (Exception, psycopg2.Error) as error:
        if (connection):
            print("Failed to insert record into mobile table", error)


try:
    path = "C:/Users/L/Desktop/uploads/"
    for e in os.listdir(path):
        if not os.path.isdir(e):
            if e.endswith('.docx'):
                docxToPdf(path + str(e), path + str(e).replace('.docx', '.pdf'))
                os.remove(path + str(e))
    for e in os.listdir(path):
        if not os.path.isdir(e):
            if e.endswith('.pdf'):
                print('******pdf')
                name = extract_name(extract_text_from_pdf(path + str(e)))
                print(name)
                email = extractEmail(extract_text_from_pdf(path + str(e)))
                print(email)
                phone = extractPhone(extract_text_from_pdf(path + str(e)))
                print(phone)
                data = ResumeParser(path + str(e)).get_extracted_data()
                skills = data['skills']
                print(skills)
                experience = extractExperience(path + str(e))
                print(experience)
                insert(name, phone, email, experience, skills)
except:
    try:
        for e in os.listdir(path):
            if not os.path.isdir(e):
                if e.endswith('.pdf'):
                    print('******pdf')
                    name = extract_name(extract_text_from_pdf(path+str(e)))
                    print(name)
                    email = extractEmail(extract_text_from_pdf(path+str(e)))
                    print(email)
                    phone = extractPhone(extract_text_from_pdf(path + str(e)))
                    print(phone)
                    data = ResumeParser(path + str(e)).get_extracted_data()
                    skills = data['skills']
                    print(skills)
                    experience = extractExperience(path+str(e))
                    print(experience)
                    insert(name, phone, email,experience, skills)
    except:
        print()
if (connection):
            cursor.close()
            connection.close()
            print("PostgreSQL connection is closed")