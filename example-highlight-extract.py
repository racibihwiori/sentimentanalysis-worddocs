#!usr/bin/python
# -*- coding: utf-8 -*-
from docx import Document
from matplotlib import pyplot as plt
import numpy as np
document = Document('Interview_with_Ricardo_Astiazaran-en-US.docx')
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
                text_element = word.find(tag_t)
                if text_element is not None and text_element.text is not None:
                    neg_value += len(text_element.text.encode('utf-8'))

            elif hi.attrib[tag_val] == 'green':
                #print (word.find(tag_t).text.encode('utf-8').lower())
                text_element = word.find(tag_t)
                if text_element is not None and text_element.text is not None:
                    pos_value += len(text_element.text.encode('utf-8'))

                
            elif hi.attrib[tag_val] == 'yellow':
                text_element = word.find(tag_t)
                if text_element is not None and text_element.text is not None:
                    neutral_value += len(text_element.text.encode('utf-8'))
            
plain_char = num_chars - neg_value - pos_value - neutral_value
user_nm = input('What is the name of the user interviewed?\n')
plt_title = "Sentiment Analysis for "+user_nm
user_inp = input('What would you like to name the image file?\n')
file_name = user_inp+'.png'
val_titles = ['Negative','Positive','Neutral','No Sentiment']
char_data = [neg_value,pos_value,neutral_value,plain_char]
fig, ax = plt.subplots(figsize=(11, 8))
colors = ['#b30118','#7db954', '#ffc502','#a4a9ad']

lgnd = ['Negative sentiment from the User', 'Positive sentiment from the User', 'Neutral Statements/Requests/Questions from the User','No particular sentiment from the User']
patches, texts, pcts = plt.pie(char_data,labels=val_titles,colors=colors,autopct='%1.1f%%', wedgeprops= {"edgecolor":"black", 
                     'linewidth': 2, 
                     'antialiased': True})
plt.legend(lgnd, loc="lower right")

#plt.title(plt_title, fontsize = 16)

for i, patch in enumerate(patches):
    #texts[i].set_color(patch.get_facecolor())
    plt.setp(pcts, color='black',fontsize=10,fontweight=700)
    plt.setp(texts,fontsize=12, fontweight=600)
    ax.set_title(plt_title, fontsize=18, fontweight=700)
    plt.tight_layout()

plt.savefig(file_name)

plt.show()


         
print('Negative Count:',neg_value,'\nPositive Count:',pos_value,'\nNeutral Count:',neutral_value,'\nTotal Characters:',num_chars)