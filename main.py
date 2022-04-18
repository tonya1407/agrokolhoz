from docxtpl import DocxTemplate
from csv import reader
# import csv
import codecs


def create_dataset(file_name):
    csvreader = reader(codecs.open(file_name, 'rU', 'utf-16'))
    # file.close()
    headers = next(csvreader)
    rows = []
    for row in csvreader:
        rows.append(row)
    dataset = []
    for count, value in enumerate(rows):
        dataset.append(dict(zip(headers, value)))
    return dataset


if __name__ == '__main__':
    data = create_dataset('Data123.csv')
    print(data)
    for count, value in enumerate(data):
        filename = str(count).zfill(4) + '_' + value['protocol_number'].replace('\\', '-').replace("/", '-') + '.docx'
        print(filename)
        doc = DocxTemplate('template.docx')
        doc.render(value)
        doc.save(filename)