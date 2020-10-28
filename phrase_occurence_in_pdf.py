import fitz #PyMuPDF
import pandas as pd


def get_page_number(row, pdf):
    occurence = []
    search_terms = str(row['Keyword'])
    for index in range(len(pdf)):
        page_number = index
        page = pdf[page_number]
        page_text = page.getText().replace('\n', ' ')
        if search_terms in page_text:
            occurence.append(str(page_number + 1))

        # return 'No Occurence'
    return ','.join(occurence)


pdf_filename = r'pdf_1.pdf'
pdf = fitz.open(f'{pdf_filename}')  # Read the file as doc

df = pd.read_excel('excel.xlsx', sheet_name='sheet1')
df['Occurence'] = df.apply(lambda row: get_page_number(row, pdf), axis=1)

df.to_csv('output.csv', index=False)
