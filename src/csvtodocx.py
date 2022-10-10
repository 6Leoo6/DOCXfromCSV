from io import BytesIO
from zipfile import ZipFile
from xml.dom import minidom


async def convertCSV(csv, model):
    # Read the bytes of the files uploaded
    csv_bytes = await csv.read()
    model_bytes = BytesIO(await model.read())

    # Parse the csv
    csv = csv_bytes.decode('utf-8').replace('\r', '').split('\n')
    csv = list(map(lambda x: x.split(','), csv))

    # Separate the first row (Headers) from the others
    csv_headers = csv.pop(0)
    csv_rows = csv

    zip_bytes = BytesIO()
    with ZipFile(zip_bytes, 'w') as zip:
        for row in csv_rows:
            files = []
            doc_xml = BytesIO()
            filename = row[0]

            with ZipFile(model_bytes, 'r') as doc:
                for file in doc.filelist:
                    fname = file.filename
                    if fname == 'word/document.xml':
                        doc_xml = doc.read(fname).decode('utf-8')
                        continue
                    files.append([fname, doc.read(fname)])

            
            for col, header in enumerate(csv_headers):
                print(header, col)
                # Replace the title with the corresponding value
                doc_xml = doc_xml.replace(header, row[col])
                # If it's the first line of the model and it starts with [ and ends with ] then set the filename to it

            xml = minidom.parseString(doc_xml)
            body = xml.firstChild.firstChild
            first_p = ''
            try:
                for child in body.firstChild.childNodes:
                    first_p += child.firstChild.firstChild.nodeValue
                if first_p.startswith('[') and first_p.endswith(']'):
                    # Remove the square brackets from both ends
                    filename = first_p.strip('[]')
                    body.removeChild(body.firstChild)
            except (AttributeError, TypeError):
                pass
            doc_xml = xml.toxml()

            doc_bytes = BytesIO()
            with ZipFile(doc_bytes, 'w') as doc:
                for fname, file in files:
                    if fname != 'word/document.xml':
                        doc.writestr(fname, file)
                doc.writestr('word/document.xml', doc_xml)

            zip.writestr(f'{filename}.docx', doc_bytes.getvalue())

    # Return the bytes of the zip
    return zip_bytes.getvalue()