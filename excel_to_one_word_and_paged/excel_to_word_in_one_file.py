import pandas as pd
from docxtpl import DocxTemplate


def read_excel_data(filepath, items_per_page=6):
    df = pd.read_excel(filepath)
    data_chunks = [df[i:i + items_per_page].to_dict(orient='records') for i in range(0, len(df), items_per_page)]
    return data_chunks


def create_word_document(template_path, data_chunks, output_path):
    doc = DocxTemplate(template_path)
    context = {'pages': data_chunks}
    doc.render(context)
    doc.save(output_path)


def main():
    excel_filepath = 'source.xlsx'
    template_path = 'template.docx'
    output_path = 'output.docx'
    items_per_page = 6

    data_chunks = read_excel_data(excel_filepath, items_per_page)
    create_word_document(template_path, data_chunks, output_path)


if __name__ == '__main__':
    main()
