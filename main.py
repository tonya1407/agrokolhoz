from codecs import open
from csv import reader
from docxtpl import DocxTemplate
from tqdm import tqdm


def create_dataset(file_name):
    csvreader = reader(open(file_name, 'rU', 'utf-16'))
    headers = next(csvreader)
    rows = []
    for row in csvreader:
        rows.append(row)
    dataset = []
    for count, value in enumerate(rows):
        dataset.append(dict(zip(headers, value)))
    return dataset


def create_documents(data, file_template, key_header, file_extention):
    filenames = []
    for count, value in enumerate(tqdm(data)):
        filename = str(count).zfill(4) + '_' + \
            value[key_header].replace('\\', '-').replace("/", '-') + \
            file_extention
        filenames.append(filename)
        doc = DocxTemplate(file_template)
        doc.render(value)
        doc.save(filename)
    print("\n".join(filenames))


if __name__ == '__main__':
    data_set = create_dataset('Data.csv')
    create_documents(data_set, 'template.docx', 'protocol_number', '.docx')
