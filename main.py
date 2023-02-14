from __future__ import print_function
from googleapiclient.errors import HttpError
import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from fpdf import FPDF
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import plotly
import os
from datetime import date


SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
ALLOWED_EXPENSES_TAB = 'Allowed expenses'
MONTH = 'February'

SAMPLE_SPREADSHEET_ID = '1TOxlVX_JRae1LC3_dMxv4nmeo7gnC3gLvBEe-UV8BA8'

SIDE_MARGIN = 10
IMG_WIDTH = 700
IMG_HEIGHT = 450


class PDF(FPDF):
    def set_title(self, title):
        self.set_font('Helvetica', 'B', 24)
        self.cell(w=0, h=0, txt=title,
                  align='C', new_x='LMARGIN', new_y='NEXT')

    def set_subtitle(self, subtitle):
        self.set_font_size(16)
        self.cell(w=0, h=24, txt=subtitle,
                  align='C', new_x='LMARGIN', new_y='NEXT')

    def add_chart(self, chart):
        self.set_x(40)
        self.image(chart,  link='', type='', w=IMG_WIDTH/5, h=IMG_HEIGHT/5)

    def add_paragraph(self, text):
        self.set_x(10)
        self.set_font('Helvetica', 'B', 16)
        self.cell(w=0, align='L', txt=text)
        self.ln(10)

    def add_text(self, text, align):
        self.set_x(10)
        self.set_font(family='Helvetica', size=12)
        self.cell(w=0, align=align, txt=text)

    def add_table(self, headings, rows):
        self.set_x(10)
        self.set_font(family='Helvetica', style='B', size=12)
        column_width = (self.w - SIDE_MARGIN * 2) / len(headings)
        for heading in headings:
            self.cell(w=column_width, h=9, txt=heading, border=1)
        self.ln()
        self.set_font(family='Helvetica', size=12)
        for row in rows:
            for col in row:
                self.cell(w=column_width, h=9, txt=col,
                                border=1)
            self.ln()


def get_parsed_sheet():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        return sheet
    except HttpError as err:
        print(err)


def get_category_expenses(month_expenses, categories):
    expenses_dictionary = {}

    for category in np.array(categories).flatten():
        expenses_dictionary[category] = 0

    for row in month_expenses:
        expense_amount = row[1]
        expense_category = row[2]
        expenses_dictionary[expense_category] += float(expense_amount or 0)

    for expense in expenses_dictionary:
        expenses_dictionary[expense] = str(
            round(expenses_dictionary[expense], 2))

    return expenses_dictionary


def generate_expenses_by_category_chart(expenses_by_category, categories, file_name='chart', file_format='png'):
    pie = px.pie(values=expenses_by_category,
                 names=categories,
                 color_discrete_sequence=px.colors.sequential.RdBu)
    plotly.io.write_image(pie, file=f'{file_name}.{file_format}',
                          format=file_format, width=IMG_WIDTH, height=IMG_HEIGHT)
    return (f'{os.getcwd()}/{file_name}.{file_format}')


def generate_expenses_by_date_chart(month_expenses, file_name='chart_2', file_format='png'):
    expenses_by_date = {}
    for expense in month_expenses:
        date = expense[0]
        amount = float(expense[1])
        if date in expenses_by_date:
            expenses_by_date[date] += amount
        else:
            expenses_by_date[date] = amount
    chart = go.Figure(data=go.Scatter(
        x=list(expenses_by_date.keys()), y=list(expenses_by_date.values()), line_shape='spline'))
    chart.update_layout(xaxis_title='Date', yaxis_title='Amount (eur.)')
    plotly.io.write_image(chart, file=f'{file_name}.{file_format}',
                          format=file_format, width=IMG_WIDTH, height=IMG_HEIGHT)
    return (f'{os.getcwd()}/{file_name}.{file_format}')

# def modify_array_element(modify_index, )


def generate_pdf():
    sheet = get_parsed_sheet()
    month_expenses = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                        range=f'{MONTH}!A3:D').execute().get('values', [])
    categories = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=f'{ALLOWED_EXPENSES_TAB}!A1:A9').execute().get('values', [])

    expenses_by_category = get_category_expenses(month_expenses, categories)

    # pdf settings
    pdf = PDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # adding title
    pdf.set_title('Monthly expenses report')

    # adding subtitle
    total_expenses = round(sum([float(i[1]) for i in month_expenses]), 2)
    pdf.set_subtitle(f'({MONTH}, totally spent: {total_expenses} eur.)')

    # adding paragraph
    pdf.add_paragraph('Expenses by category:')

    # adding chart
    chart = generate_expenses_by_category_chart(
        expenses_by_category, categories)
    pdf.add_chart(chart)

    # adding expenses by category table
    pdf.add_table(('Category', 'Expense amount'),
                  [(i[0], f'{i[1]} eur.') for i in list(expenses_by_category.items())])

    # adding a new page
    pdf.add_page()
    pdf.add_paragraph('Expenses by date:')

    # adding chart
    chart = generate_expenses_by_date_chart(month_expenses)
    pdf.add_chart(chart)
    pdf.ln(10)

    # adding total expenses table
    total_expenses_headers = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                                  range=f'{MONTH}!A2:C2').execute().get('values', [])[0]
    
    pdf.add_table(total_expenses_headers,
                  [(i[0], f'{i[1]} eur.', *i[2:-1]) for i in month_expenses])

    # write a pdf
    pdf.output(f'{MONTH}_{date.today()}.pdf')


def main():
    generate_pdf()


if __name__ == '__main__':
    main()
