# -*- coding: utf-8 -*-
from email import encoders
from email.header import Header
from email.mime.text import MIMEText
from email.utils import parseaddr, formataddr
import smtplib
import pandas as pd

class AutoMailer():
    def __init__(self, conf, database, template, format_path, map_path=None):
        self.conf_path = conf
        self.database = database
        self.mail_data = pd.DataFrame()
        self.template = template
        self.format_path = format_path
        self.format_info = []
        self.map_path = map_path
        self.mail_body = None

    def __del__(self):
        self.server.quit()

    def __format_name(self, name):
        return formataddr((Header(name, 'utf-8').encode(), '')).split('<')[0]

    def __readConf(self):
        conf = {}
        with open(self.conf_path) as fp:
            lines = fp.read().splitlines()
        for i in range(len(lines)):
            key, val = lines[i].split(':', 1)
            conf[key] = val
        self.header = conf['header']
        self.host = conf['host']
        self.hostname = conf['hostname']
        self.pwd = conf['pwd']
        self.test_addr = conf['test_addr']
        self.smtp_server = conf['smtp_server']
        self.smtp_port = conf['smtp_port']
        self.server = smtplib.SMTP()
        self.__printLog('host: ' + self.host)
        self.__printLog('pwd: ' + self.pwd)
        self.__printLog('test_addr: ' + self.test_addr)
    
    def __initSmtp(self):
        self.server = smtplib.SMTP(self.smtp_server, self.smtp_port)
        self.server.starttls()
        self.server.login(self.host, self.pwd)

    def __readSendData(self):
        '''
            Read the google form's data.
        '''
        self.mail_data = pd.read_csv(self.database, dtype=str)
        self.mail_data = self.mail_data.fillna('')
        if self.map_path is not None:
            map_info = {}
            with open(self.map_path) as fp:
                lines = fp.read().splitlines()
            for line in lines:
                key, val = line.split(':', 1)
                map_info[key] = val
            self.mail_data = self.mail_data.rename(columns=map_info)

    def __getFormat(self, val):
        '''
            Read the reply mail format
            TODO:
                - [] Calculate the fee, or do it on the registry google form
        '''
        with open(self.format_path) as fp:
            keys = fp.read().split(',')
        for i in range(len(keys)-3):
            self.format_info.append(val[keys[i]])
        self.format_info.append('$')
        self.format_info.append('$')
        self.format_info.append('$')

    def __setMailBody(self, val):
        '''
			Set the email body in html format.

			Attribute:
				val: A datafram's row, which need to contain 'to_name' and 'to_addr' two columns
        '''
        with open(self.template) as fp:
            contents = fp.read()
        self.format_info = []
        self.__getFormat(val)
        self.mail_body = contents.format(*self.format_info)
        self.mail_body = MIMEText(self.mail_body, 'html', 'utf-8')
        self.mail_body['From'] = f'{self.hostname} <{self.host}>'
        self.mail_body['To'] = self.__format_name(val['to_name'])+ f' <{val["to_addr"]}>'
        self.mail_body['Subject'] = Header(self.header, 'utf-8').encode()

    def __sendMail(self):
        self.__printLog('<Sending start>')
        for _, row in self.mail_data.iterrows():
            self.__setMailBody(row)
            try:
                self.__printLog('To ' + row['to_name'] + ': ' + row['to_addr'])
                self.server.sendmail(self.host, self.test_addr, self.mail_body.as_string())
                self.__printLog('\tsent!')
            except smtplib.SMTPException:
                self.__printLog('\tsent fail!')
        self.__printLog('<Sending finish>')

    def __printLog(self, msg):
        print('[LOG] '+str(msg))

    def work(self):
        self.__readConf()
        self.__initSmtp()
        self.__readSendData()
        self.__sendMail()

def main():
    print('[LOG] Debugging mode')
    CONFIG = '../config.txt'
    DATABASE = '../database/test_data.csv'
    MAP_INFO = '../template/test_mapping.txt'
    TEMPLETE = '../template/test_templete.html'
    FORMAT_INFO = '../template/test_format.fmt'

    mailer = AutoMailer(CONFIG, DATABASE, TEMPLETE, FORMAT_INFO, MAP_INFO)
    mailer.work()

if __name__ == "__main__":
    main()
