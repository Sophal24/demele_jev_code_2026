import requests
import pandas as pd
from io import StringIO
import json
import os

# Column name Mapping - Testing
df = pd.read_excel('../codebook_demele_cohort1.xlsx', sheet_name='Baseline')
df = df[['Question','Response','Code','Remark']]

df = df[['Code', 'Question']]
df = df.dropna(how='all')
df.to_excel('../result/question_code_mapping.xlsx', index=False)


# Form ID Setting
url_a1 = "A1_Questionnaire_Cohort1_Baseline"
url_a2 = "A2_Questionnaire_Cohort1_Y1"
url_a3 = "A3_Questionnaire_Cohort1_Y2"
url_a4 = "A4_Questionnaire_Cohort2_Baseline"
url_a5 = "A5_Questionnaire_Cohort2_M1"
url_a6 = "A6_Questionnaire_Cohort2_M3"
url_a7 = "A7_Questionnaire_Cohort2_M6"
url_a8 = "A8_Questionnaire_Cohort2_Y1"
url_b1 = "B1_Sample_Collection_Form_Cohort1_Baseline"
url_b2 = "B2_Sample_Collection_Form_Cohort1_Y1"
url_b3 = "B3_Sample_Collection_Form_Cohort1_Y2"
url_b4 = "B4_Sample_Collection_Form_Cohort2_Baseline"
url_b5 = "B5_Sample_Collection_Form_Cohort2_M1"
url_b6 = "B6_Sample_Collection_Form_Cohort2_M3"
url_b7 = "B7_Sample_Collection_Form_Cohort2_M6"
url_b8 = "B8_Sample_Collection_Form_Cohort2_Y1"
url_c1 = "C1_LabResult_Cohort1_Baseline_ELISA"
url_c2 = "C2_LabResult_Cohort1_Y1_ELISA"
url_c3 = "C3_LabResult_Cohort1_Y2_ELISA"
url_c4 = "C4_LabResult_Cohort2_Baseline_ELISA"
url_c5 = "C5_LabResult_Cohort2_Baseline_QIASTAT_PCR"
url_c6 = "C6_LabResult_Cohort2_Baseline_ IPC_GenXpert"
url_c7 = "C7_LabResult_Cohort2_Baseline_IPC_TB_Culture"
url_c8 = "C8_LabResult_Cohort2_M1_ELISA"
url_d1 = "D1_Screening_Form_Cohort1_Baseline"
url_d2 = "D2_Screening_Form_Cohort1_Y1"
url_d3 = "D3_Screening_Form_Cohort1_Y2"
url_d4 = "D4_Screening_Form_Cohort2_Baseline"
url_d5 = "D5_Screening_Form_Cohort2_M1"
url_d6 = "D6_Screening_Form_Cohort2_M3"
url_d7 = "D7_Screening_Form_Cohort2_M6"
url_d8 = "D8_Screening_Form_Cohort2_Y1"
url_c9a = "C9a_LabResult_PDF_IPC_GenXpert"
url_c9b = "C9b_LabResult_PDF_IPC_TB_Culture"


# Function - Get data from Form ID URL

def get_data_url(form_id):
    params = {
        "splitSelectMultiples": "true",
        "groupPaths": "false",
        "deletedFields": "true"
    }
    url = 'https://demele-jev.uhs.edu.kh/api/download-submissions-v3?formId=' + form_id + '&secret_key=13ac3c2f-b7df-4ce9-8010-247da09a91d9'
    
    response = requests.get(url, params = params, timeout=120)
    # response = requests.get(url)
    
    df = pd.read_csv(StringIO(response.text))
    return df

def check_basic_info(df):
    df = get_data_url(url_a1)
    print(df.columns)
    print(df.shape)
    print(df.info())