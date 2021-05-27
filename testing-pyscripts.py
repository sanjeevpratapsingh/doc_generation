import os.path

from docx import Document
from docx.shared import Pt,Cm
from docx.enum.style import WD_STYLE_TYPE
from docxtpl import DocxTemplate
import re

phone = "2004-959-559# This is Phone Number? Hi /"

# Delete Python-style comments
num = re.sub(r'[^a-zA-Z0-9 \n\.]',"", phone)
print("Phone Num : ", num)

# Remove anything other than digits
num = re.sub(r'\D', "", phone)    
print("Phone Num : ", num)


my_str = "hey th~!ere"
my_new_string = re.sub(r'[^a-zA-Z0-9 \n\.]', '', my_str)
print(my_new_string)



doc = DocxTemplate("demodata.docx")

doc.save("generated_doc.docx")

# document  = Document()
# styles = document.styles
# font = document.styles['Normal'].font
# font.name = "Poppins"
# font.size = Pt(14)

# heading_style = document.styles['Heading 2'].font
# heading_style.name = "Poppins SemiBold"
# heading_style.size = Pt(40)

# document.add_heading("This a Test Doc.")

# sections = document.sections
# for section in sections:
#     margin_size = 2
#     section.left_margin = Cm(margin_size)
#     section.right_margin = Cm(margin_size)
#     section.top_margin = Cm(margin_size)
#     section.bottom_margin = Cm(margin_size)


# filepath = 'C:/Users/Sanjeev Pratap Singh/Desktop/'

# document.save(filepath + 'demo.docx')
