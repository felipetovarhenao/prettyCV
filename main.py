from cv import CV
from datetime import date
import sys

cv = CV()
cv_path = '../personal-website/src/json/cv.json'
works_path = '../personal-website/src/json/work-catalog.json'
cv.load_data(cv_path=cv_path, works_path=works_path)
cv.compile()
file_id = date.today()
file = 'cv'
file_doc = file + '.docx'
if len(sys.argv) > 1 and sys.argv[1] == '--local':
    cv.write(file_doc, open_file=False)
else:
    cv.write(f'/Users/felipe-tovar-henao/Google Drive/My Drive/FTH Drive/CV/CV_{file_id}.docx')
