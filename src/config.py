import os
import configparser

config = configparser.ConfigParser()
config.read('./config.ini')

DEBUG = config['settings']['debug'] == 'True'
DOCUMENT_PATH = os.path.abspath(os.path.join(os.getcwd(), "documents"))
DATA_PATH = os.path.abspath(os.path.join(os.getcwd(), "data"))
TEMPLATE_DIR_PATH = os.path.abspath(os.path.join(os.getcwd(), "templates"))
TEMPLATE_PATH = os.path.abspath(os.path.join(TEMPLATE_DIR_PATH, "berichtsheft_empty.docx"))

NO_WORK_DAY = {
    'from': config['no_work_day']['from'],
    'to': config['no_work_day']['to'],
    'is': config['no_work_day']['is']
}
SCHOOL_DAYS = [
    {
        'from': config['school_0']['from'],
        'to': config['school_0']['to'],
        'is': config['school_0']['is']
    },
    {
        'from': config['school_1']['from'],
        'to': config['school_1']['to'],
        'is': config['school_1']['is']
    },
    {
        'from': config['school_2']['from'],
        'to': config['school_2']['to'],
        'is': config['school_2']['is']
    },
    {
        'from': config['school_3']['from'],
        'to': config['school_3']['to'],
        'is': config['school_3']['is']
    },
    {
        'from': config['school_4']['from'],
        'to': config['school_4']['to'],
        'is': config['school_4']['is']
    },
    NO_WORK_DAY,
    NO_WORK_DAY
]
WORK_DAY = {
    'from': config['work_day']['from'],
    'to': config['work_day']['to'],
    'is': config['work_day']['is']
}
STATICS = {
    'ausbildungsbereich': config['statics']['ausbildungsbereich'],
    'person': '',  # Is detected from .csv
    'whole_time': config['statics']['whole_time'],
    'school_week': None
}
SCHOOL_JOB = config['easyjob_settings']['school_job']
SCHOOL_JOB_CSV = config['easyjob_settings']['school_job_csv']
AUSBILDUNGSABSCHNITT = config['ihk_settings']['ausbildungsabschnitt']
BETREUEREMAIL = config['ihk_settings']['betreueremail']
EASYJOB_SETTINGS = {
    'username': config['easyjob_settings']['username'],
    'password': config['easyjob_settings']['password'],
    'url': config['easyjob_settings']['url']
}
IHK_SETTINGS = {
    'username': config['ihk_settings']['username'],
    'password': config['ihk_settings']['password'],
    'url': config['ihk_settings']['url']
}
