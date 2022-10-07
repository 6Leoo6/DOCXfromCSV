from docx import Document
from io import BytesIO
from zipfile import ZipFile

async def convertCSV(csv, model):
    csv_bytes = await csv.read()
    model_bytes = BytesIO(await model.read())

    csv = csv_bytes.decode('utf-8').replace('\r', '').split('\n')
    csv = list(map(lambda x: x.split(','), csv))

    headers = csv.pop(0)
    rows = csv
    
    print(headers, rows)

    doc = Document(model_bytes)
    zip_bytes = BytesIO()
    with ZipFile(zip_bytes, 'w') as zip:
        for csv_row in rows:
            fullDocument = []
            filename = csv_row[0]
            for i, paragraph in enumerate(doc.paragraphs):
                paragraph = paragraph.text
                for col, header in enumerate(headers):
                    paragraph = paragraph.replace('${'+header+'}', csv_row[col])
                if i==0 and paragraph.startswith('[') and paragraph.endswith(']'):
                    filename = paragraph.strip('[]')
                    continue
                fullDocument.append(paragraph)

            generatedDoc = Document()
            for paragraph in fullDocument:
                generatedDoc.add_paragraph(paragraph)
            doc_bytes = BytesIO()
            generatedDoc.save(doc_bytes)
            zip.writestr(f'{filename}.docx', doc_bytes.getvalue())
    
    return zip_bytes.getvalue()