import pandas as pd
from pandas import DataFrame, read_csv, to_datetime
from datetime import datetime, timedelta
import copy
import shutil
from docxtpl import DocxTemplate
from src.config import *
from src.browser import setupBrowser
import os
from sys import exit

from src.utils import printProgressBar, get_date_data

leistung_key = 'Leistung'


class Manager:

    def createIhkEntrys(self):
            browser = setupBrowser()
            browser.createWeekEntrys(self.weeks)

    def createDocuments(self):
        print("Starting with Document generation")
        printProgressBar(0, len(self.weeks), prefix='Processing Document:', suffix='Complete', length=50)
        for idx, week in enumerate(self.weeks):
            # Copy word Document
            destination_path = os.path.join(DOCUMENT_PATH, str(week['days'][0]['date']) + ".docx")
            shutil.copyfile(TEMPLATE_PATH, destination_path)
            abs_destination_path = os.path.abspath(destination_path)

            doc_context = {
                'person': week['static']['person'],
                'year': week['static']['year'],
                'ausbildungsbereich': week['static']['ausbildungsbereich'],
                'week_from': week['days'][0]['date'],
                'week_to': week['days'][6]['date'],
                'whole_hours': week['static']['whole_time'],
            }

            for i in range(0, 7):
                doc_context[f'day_from{i}'] = week['days'][i]['time']['from']
                doc_context[f'day_to{i}'] = week['days'][i]['time']['to']
                doc_context[f'hours{i}'] = week['days'][i]['time']['is']
                doc_context[f'task{i}'] = '\n'.join(week['days'][i]['descriptions'])

            template = DocxTemplate(abs_destination_path)
            template.render(doc_context)
            template.save(abs_destination_path)
            printProgressBar(idx, len(self.weeks), prefix='Processing Document:', suffix='Complete', length=50)

    def start(self):
        self.df = self.get_dataframe()
        self.process_person()
        self.process_days()
        self.process_weeks()
        if IHK_SETTINGS['to_ihk']:
            self.createIhkEntrys()
        self.createDocuments()
        print("Done")

    def process_person(self):
        person = self.df['Mitarbeiter'].iat[0]
        STATICS['person'] = person

    def process_days(self):
        # Get all dates in the table
        self.min_date = self.df['Datum'].min().to_pydatetime()
        self.max_date = self.df['Datum'].max().to_pydatetime()
        # Process the dates
        # Check if last day is a friday
        week_day = self.min_date.weekday()
        if week_day > 0:
            self.min_date += timedelta(days=7 - week_day)

        # Check if max user input is a sunday
        week_day = self.max_date.weekday()
        if week_day < 4:
            self.max_date -= timedelta(days=week_day)
        elif week_day == 4: # Set max Date to sunday if it is friday
            self.max_date += timedelta(days=2)

        delta = (self.max_date - self.min_date)
        self.days = delta.days

    def process_weeks(self):
        weeks = []
        week = {
            'static': copy.deepcopy(STATICS),
            'days': []
        }
        print("Processing Weeks:")
        for i in range(self.days + 1):
            day = self.min_date + timedelta(days=i)

            process_date = day.strftime('%Y.%m.%d')
            day_data = dict({
                'date': day.strftime('%d.%m.%Y')
            })
            if week['static']['school_week'] is None:
                week['static']['year'] = day.strftime('%Y')
                print("Week: " + str(len(weeks) + 1))
                job_values = self.df[self.df['Datum'] == pd.Timestamp(day)][leistung_key].unique()
                school_day = SCHOOL_JOB in job_values
                if school_day:
                    week['static']['school_week'] = True
                else:
                    week['static']['school_week'] = False

            if week['static']['school_week']:
                day_data['time'] = SCHOOL_DAYS[len(week['days'])]
            else:
                day_data['time'] = WORK_DAY if len(week['days']) < 5 else NO_WORK_DAY  # Check witch Weekday
            descriptions = self.df[self.df['Datum'] == process_date]['Beschreibung'].to_numpy().tolist() if len(week['days']) < 5 else ['frei']
            day_data['descriptions'] = descriptions if len(descriptions) > 0 else ['frei']
            week['days'].append(day_data)
            if len(week['days']) > 6:
                if week['static']['school_week']:
                    week['static']['whole_time'] = '29:00'
                weeks.append(week)
                week = {
                    'static': copy.deepcopy(STATICS),
                    'days': []
                }

        self.weeks = weeks

    def get_dataframe(self) -> DataFrame:
        """
        returns the first csv in ./data as DataFrame
        :return: pd.DataFrame
        """
        if EASYJOB_SETTINGS['from_ej']:
            browser = setupBrowser()
            return browser.processEJ()

        global leistung_key
        leistung_key = "Art, Zuordnung Leistung Beschreibung"
        for file in os.listdir(DATA_PATH):
            # Check for .csv files in data dir
            if not file.endswith('.csv'):
                continue

            file_path = os.path.join(DATA_PATH, file)

            # Create dataframe from csv and return
            df = read_csv(file_path, sep=";")
            df["Datum"] = to_datetime(df["Datum"])
            df.sort_values(by='Datum', inplace=True, ascending=False)
            print("Data imported")
            return df
        exit("No .csv File found")

    def get_date_input(self, type):
        while True:
            try:
                user_input = input(f"{type} Date (Format 20.12.2022):")
                datetime.strptime(user_input, '%d.%m.%Y')
                EASYJOB_SETTINGS[f'{type}_date'] = user_input
                break
            except ValueError:
                print("Format not valid")

    def day_of_week_num(self, dts):
        return (dts.astype('datetime64[D]').view('int64') - 4) % 7

    def check_easyjob_settings(self):
        if input("Chose Dates automatic? y/n (default=y): ") in ["y", ""]:
            browser = setupBrowser()
            EASYJOB_SETTINGS['from_date'] = browser.getNeededIhkDates().strftime('%d.%m.%Y')
            EASYJOB_SETTINGS['to_date'] = datetime.now().strftime('%d.%m.%Y')
        else:
            self.get_date_input("from")
            self.get_date_input("to")

        delta, min_date, max_date = get_date_data(EASYJOB_SETTINGS['from_date'], EASYJOB_SETTINGS['to_date'])
        if delta.days < 5:
            exit("Weeks already processed!")


    def __init__(self):
        # Add defaults for needed fields
        self.weeks = None
        self.max_date = None
        self.min_date = None
        self.days = None
        self.df = None

        dirs = [TEMPLATE_DIR_PATH, DOCUMENT_PATH, DATA_PATH]
        for dir in dirs:
            if not os.path.exists(dir):
                os.makedirs(dir)
        if not os.path.exists(TEMPLATE_PATH):
            exit(f"No Template File | Path: {TEMPLATE_PATH}")
        if input("Do you want to post directly to ihk? y/n (default=y): ") in ["y", ""]:
            IHK_SETTINGS['to_ihk'] = True
        else:
            IHK_SETTINGS['to_ihk'] = False
        if input("Do you want to pull from easyjob? y/n (default=y): ") in ["y", ""]:
            EASYJOB_SETTINGS['from_ej'] = True
            self.check_easyjob_settings()
        else:
            EASYJOB_SETTINGS['from_ej'] = False
            SCHOOL_JOB = SCHOOL_JOB_CSV
            print(f"Put your easyJob export as .csv in {os.path.abspath(DATA_PATH)}")
            input("Press any key to proceed")
