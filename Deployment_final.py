import os
import re
import PyPDF2
import docx2txt
from PyPDF2 import PdfReader
import pandas as pd
import streamlit as st
import win32com.client
import joblib
import textract

import en_core_web_sm
nlp = en_core_web_sm.load()
import nltk
from nltk.corpus import stopwords
 
nltk.download('stopwords')
nltk.download('wordnet')

from nltk.stem import WordNetLemmatizer
from nltk.tokenize import RegexpTokenizer

import pythoncom

pythoncom.CoInitialize()

#----------------------------------------------------------------------------------------------------

st.title('RESUME CLASSIFICATION APP')
st.markdown('<style>h1{color: Purple;}</style>', unsafe_allow_html=True)


def extract_skills(resume_text):
    nlp_text = nlp(resume_text)
    noun_chunks = nlp_text.noun_chunks
    tokens = [token.text for token in nlp_text if not token.is_stop]

    data = pd.read_csv(r"skills.csv") 
    skills = list(data.columns.values)
    skillset = []

    for token in tokens:
        if token.lower() in skills:
            skillset.append(token)

    for token in noun_chunks:
        token = token.text.lower().strip()
        if token in skills:
            skillset.append(token)   
    return [i.capitalize() for i in set([i.lower() for i in skillset])]

def getText(resume):
    app = win32com.client.Dispatch('Word.Application')
    resume_data = []
    directory = f'{resume}'
    if str(resume.name).endswith('.docx'):
        resume_data.append((textract.process(f'{os.getcwd()}\\Test\\{resume.name}').decode('utf-8')))
    elif str(resume.name).endswith('.doc'):
        doc = app.Documents.Open(f'{os.getcwd()}\\Test\\{resume.name}')
        doctext = doc.Content.Text
        doc.Close()
        resume_data.append(doctext)
    elif str(resume.name).endswith('.pdf'):
        reader = PdfReader(f'{os.getcwd()}\\Test\\{resume.name}')
        text = ""
        for page in reader.pages:
            text += page.extract_text() + "\n"
        resume_data.append(text)
    app.Quit()
    return resume_data


def preprocess(sentence):
    sentence = str(sentence)
    sentence = sentence.lower()
    sentence = sentence.replace('{html}',"") 
    cleanr = re.compile('<.*?>')
    cleantext = re.sub(cleanr, '', sentence)
    rem_url = re.sub(r'http\S+', '',cleantext)
    rem_num = re.sub('[0-9]+', '', rem_url)
    tokenizer = RegexpTokenizer(r'\w+')
    tokens = tokenizer.tokenize(rem_num)  
    filtered_words = [w for w in tokens if len(w) > 2 if not w in stopwords.words('english')]
    lemmatizer = WordNetLemmatizer()
    lemma_words = [lemmatizer.lemmatize(w) for w in filtered_words]
    return " ".join(lemma_words) 

file_type=pd.DataFrame([], columns=['Resume', 'Predicted Profile', 'Skills',])
filename = []
predicted = []
skills = []


import pickle as pk
model = joblib.load('random_forest.pkl')
Vectorizer = joblib.load('unigram_tfidf.pkl')

upload_file = st.file_uploader('Upload your resumes here', type= ['docx', 'doc', 'pdf'],accept_multiple_files=True)
  
        
select = ['PeopleSoft Developer','SQL Developer','React JS Developer','Work_Day Developer']
user_choice = st.multiselect("Select the categories of resumes you want to see. To see all leave this unselected.", options=select)

if st.button('Predict'):
    for doc_file in upload_file:
        if doc_file is not None:
            filename.append(doc_file.name)
            extText = getText(doc_file)
            cleaned = preprocess(extText)
            prediction = model.predict(Vectorizer.transform([cleaned]))[0]
            predicted.append(prediction)
            skills.append(extract_skills(extText[0]))
        
    if len(predicted) > 0:
        file_type['Resume'] = filename
        file_type['Skills'] = skills
        file_type['Predicted Profile'] = predicted
        if user_choice:
            st.table(file_type[file_type['Predicted Profile'].isin(user_choice)])
        else:
            st.table(file_type.style.format())
    else:
        st.write('Please upload the files first.')