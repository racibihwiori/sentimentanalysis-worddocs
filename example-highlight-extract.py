#!usr/bin/python
# -*- coding: utf-8 -*-
from docx import Document
from matplotlib import pyplot as plt
import numpy as np
document = Document('Interview_User_Name_Color_Coded.docx')
words = document._element.xpath('//w:r')
WPML_URI = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
tag_rPr = WPML_URI + 'rPr'
tag_highlight = WPML_URI + 'highlight'
tag_val = WPML_URI + 'val'
tag_t = WPML_URI + 't'
chars = [chars for paragraph in document.paragraphs for chars in paragraph.text]
num_chars =len(chars)
neg_value = 0
pos_value = 0
neutral_value = 0
for word in words:        

    for rPr in word.findall(tag_rPr):
        high=rPr.findall(tag_highlight)            

        for hi in high:

            if hi.attrib[tag_val] == 'red':
                neg_value+=len(word.find(tag_t).text.encode('utf-8'))
            elif hi.attrib[tag_val] == 'green':
                #print (word.find(tag_t).text.encode('utf-8').lower())
                pos_value+=len(word.find(tag_t).text.encode('utf-8'))

                
            elif hi.attrib[tag_val] == 'yellow':
                
                neutral_value+=len(word.find(tag_t).text.encode('utf-8'))
            
plain_char = num_chars - neg_value - pos_value - neutral_value
user_nm = input('What is the name of the user interviewed?\n')
plt_title = "Sentiment Analysis for "+user_nm
user_inp = input('What would you like to name the image file?\n')
file_name = user_inp+'.png'
val_titles = ['Negative','Positive','Neutral','Unlabeled']
char_data = [neg_value,pos_value,neutral_value,plain_char]
fig = plt.figure(figsize=(10,10))
colors = ['red','green', 'yellow','gray']
plt.pie(char_data,labels=val_titles,colors=colors,autopct='%1.1f%%')
plt.title(plt_title)
plt.savefig(file_name)
plt.show()           
print('Negative Count:',neg_value,'\nPositive Count:',pos_value,'\nNeutral Count:',neutral_value,'\nTotal Characters:',num_chars)
