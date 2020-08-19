conda install -c conda-forge docxtpl

#Load the Libraries
from docxtpl import DocxTemplate
import jinja2
import pandas as pd

#Load in our template for the FS, replace note 6 with your template of choice
doc = DocxTemplate('note6 - template.docx')
context = {
    'cur_year' : 2020,
    'prior_year' : 2019
}
doc.render(context)
doc.save('note_output1.docx')

#Load in our note data from the excel file, replace notes_data with your data file
note_data = pd.read_excel('notes_data.xlsx')

#Creating a context dictionary from note data file
context = dict(zip(note_data['var'], note_data['value']))

#Format our numbers so that they have commas
def comma(value):
    return "{:,}.format(value)"
jinja_env = jinja2.Environment()
jinja_env.filters['c'] = comma

#Create our output file with all of our updates, change note 6 with your template and change save to your desired output
doc = DocxTemplate('note6 - template.docx')
doc.render(context)
doc.save('note_output.docx')




