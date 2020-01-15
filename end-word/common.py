# Standard imports
import os, datetime

# Third-party imports
import docx  # To read docx and extract data

# Get paths to sample content and data
sample_path = os.path.join(os.curdir,'samples')
contents = []  # Initialise an empty list to store paths to contents files

for dirname, dirnames, filenames in os.walk(sample_path):
    for fname in filenames:
        if 'content' in fname:
            contents.append(os.path.join(dirname,fname))
        elif 'backpage' in fname:
            backpage = os.path.join(dirname,fname)

# Sort contents list
contents.sort()
            
# Create empty output file as placeholder
output = docx.Document()
output_path = os.path.join(sample_path,'0 output.docx')
output.save(output_path)

# Create temp folder. Folder deleted after publishing
# Use tempfile when code is rebased
temp_path = os.path.join(os.curdir,'temp')
try:
    os.mkdir(temp_path)
except FileExistsError as e:
    print(f'{e}: Folder already exists...continuing\n')
    pass

# Initiate template path to title page and content to fill title+backpage
title_page = os.path.join(sample_path,'1 template.docx')
context = {
    'title': 'Prototyping with Bob',
    'subtitle': 'Prepared by Yemeng Bob Jin for Yeqin Jim Jin',
    'date': datetime.date.today(),
    'closing': 'THANK YOU',
    'copyright': 'Give me a shout out and you can do whatever (GNU Licence)',
    'website': 'www.bobjin.me',
    'email': 'automaticjinandtonic@gmail.com',
    'number': '+61 4XX XXX XXX'
}