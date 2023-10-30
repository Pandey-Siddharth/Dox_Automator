from docx import Document
import pandas as pd

def fill_invitation(template_path, output_path, data):
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        for key,value in data.items():
            if key in paragraph.text:
                for run in paragraph.runs:
                    run.text = run.text.replace(key,value)

    doc.save(output_path)


def fill_invitation_automatic(csv_path, template_path):
    df = pd.read_csv(csv_path)
    for idx, rows in df.iterrows():
        data = {
            '[Your Name]': rows['Your Name'],
            #'[Age]': rows[' Age'],
            '[Date of the Party]': rows[' Date of the Party'],
            '[Start Time]': rows[' Start Time'],
            '[Party Venue or Address]': rows[' Party Venue or Address'],
            '[Friends Name / Guests Name]': rows[' Friends Name / Guests Name'],
            '[Theme]': rows[' Theme'],
            '[RSVP Deadline]': rows[' RSVP Deadline'],
            '[Your Contact Information]': rows[' Your Contact Information']
        }
        output_path = f'Invitation_{idx+1}.docx'
        fill_invitation(template_path,output_path,data)




if __name__ == '__main__':
    csv_path = 'entries.csv'
    template_path = 'Template.docx'
    fill_invitation_automatic(csv_path,template_path)


