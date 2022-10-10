from io import BytesIO
from zipfile import ZipFile
from xml.dom import minidom
from numtoletter import int_to_letter


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

    #Create a new empty zip in memory
    zip_bytes = BytesIO()
    #Open it for writing
    with ZipFile(zip_bytes, 'w') as zip:
        #Loop through the rows of the csv
        for row in csv_rows:
            files = []
            #Create an empty XML file in memory
            doc_xml = BytesIO()
            #Define the default filename
            filename = row[0]
            
            #Open the document model as a zip
            with ZipFile(model_bytes, 'r') as doc:
                #Loop through the files of the document
                for file in doc.filelist:
                    fname = file.filename
                    #If the file is the main file then save it for later use and continue
                    if fname == 'word/document.xml':
                        doc_xml = doc.read(fname).decode('utf-8')
                        continue
                    #Store all the other files temporarly to prevent creating duplicates
                    files.append([fname, doc.read(fname)])

            #Loop through the values in the csv row
            #Reverse the row to prevent the replacement of VARAB by VARA
            #(Loop in descending order by the length of the id's)
            for col, _ in reversed(list(enumerate(csv_headers))):
                # Replace the id with the corresponding value
                current_id = f'XVAR{int_to_letter(col+1)}'
                doc_xml = doc_xml.replace(current_id, row[col])

            #Parse the replaced XML
            xml = minidom.parseString(doc_xml)
            #Get the body of the file
            body = xml.firstChild.firstChild
            first_p = ''
            try:
                #Reconstruct the first paragraph by looping through it and adding the values
                for child in body.firstChild.childNodes:
                    first_p += child.firstChild.firstChild.nodeValue
                #If the text start with and ends with a square bracket
                if first_p.startswith('[') and first_p.endswith(']'):
                    #Set the parsed paragraph as the filename
                    filename = first_p.strip('[]')
                    #Remove the paragraph from the document
                    body.removeChild(body.firstChild)
            except (AttributeError, TypeError): #If the first element is an image, there will be an error
                pass
            #Convert back to text
            doc_xml = xml.toxml()

            #Create an empty doc
            doc_bytes = BytesIO()
            #Read the doc as a zip
            with ZipFile(doc_bytes, 'w') as doc:
                #Put back the files to the zip
                for fname, file in files:
                    if fname != 'word/document.xml':
                        doc.writestr(fname, file)
                #Put back the modified XML
                doc.writestr('word/document.xml', doc_xml)

            #Save the file as a docx to the main zip
            zip.writestr(f'{filename}.docx', doc_bytes.getvalue())

    # Return the bytes of the main zip
    return zip_bytes.getvalue()
