import pandas as pd
import re
from time import sleep
from itertools import chain

class clean_data():
    def __init__(self, data_objetc, nome, email_columns):
        if type(email_columns) == str:
            if (',' in email_columns) and (';' not in email_columns):
                email_columns = email_columns.split(',')
                for i, v in enumerate(email_columns):
                    email_columns[i] = v.strip()
            elif (';' in email_columns) and (',' not in email_columns):
                email_columns = email_columns.split(';')
                for i, v in enumerate(email_columns):
                    email_columns[i] = v.strip()
            else:
                email_columns = [email_columns]
        else:
            print('O valor inserido é invalido')

        self.data = data_objetc.drop_duplicates(subset=email_columns, keep='last')
        self.nome = nome
        self.email_columns = email_columns

        self.data = self.clean_names()
        self.data = self.clean_emails().drop_duplicates(subset='E-mail')

    
    def clean_names(self):
        
        self.data[self.nome] = self.data[self.nome].str.strip()
        self.data[self.nome] = self.data[self.nome].str.replace('/', '.')
        self.data[self.nome] = self.data[self.nome].apply(lambda x: re.sub(r'[0-9]', '', x))
        self.data[self.nome] = self.data[self.nome].apply(lambda x: re.sub(r'/\n', '', x))
        self.data[self.nome] = self.data[self.nome].apply(lambda x: re.sub(r'P/', '', x))
        self.data[self.nome] = self.data[self.nome].str.strip()

        return self.data

    def clean_emails(self):

        for i in self.email_columns:
            self.data[i] = self.data[i].str.strip()
            self.data[i] = self.data[i].str.lower()
        
        self.emails_list = []

        for i in self.data[self.nome].unique():
            self.emails_list.append(self.data.groupby(self.nome).get_group(i)[self.email_columns].values)

        self.df = pd.DataFrame({'Nome':self.data[self.nome].unique(),'E-mail':self.emails_list})

        self.df['E-mail'] = self.df['E-mail'].apply(lambda x: '; '.join([str(elem) for elem in x]))

        self.df['E-mail'] = self.df['E-mail'].apply(self.email_validation)

        self.df['E-mail'] = self.df['E-mail'].apply(lambda x: '; '.join([str(elem) for elem in x]))

        self.df = self.df[self.df['E-mail'].str.len() > 4]

        return self.df


    def email_validation(self, x):
        result = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', x)
        return result

if __name__ == '__main__':
    cleanner = clean_data(
        "Excel ou CSV",
        "Coluna com o identificados, geralmente o nome",
         "Colunas com os e-mails, seperardos por virgula ou ponto e virgula"
        )