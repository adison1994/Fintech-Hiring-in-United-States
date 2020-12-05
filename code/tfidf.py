import pdfminer
import re
import xlrd
import pandas as pd
import openpyxl as op
import collections
import operator
import csv
import math
from textblob import TextBlob as tb



from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import BytesIO
from collections import Counter

def convert_pdf_to_txt(path):
    resourceManager = PDFResourceManager()
    returnstream = BytesIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(resourceManager, returnstream, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(resourceManager, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password,caching=caching, check_extractable=True):
        interpreter.process_page(page)

    text = returnstream.getvalue()

    fp.close()
    device.close()
    returnstream.close()
    return (text)



ms = op.load_workbook('Stopwords.xlsx')

ws = ms.active
listStopWords=[]

for row in ws.iter_rows(min_row=1, max_col=1, max_row=1416):
    for cell in row:
        listStopWords.append(cell.value) 


# print(listStopWords)
# print(len(listStopWords))
#print(z)





files='Beyond_Fintech_-_A_Pragmatic_Assessment_of_Disruptive_Potential_in_Financial_Services.pdf'
fileToText = convert_pdf_to_txt(files)
y=fileToText
z=y.decode("utf-8") 
z=re.sub('[^A-Za-z]+', ' ', z)
zFinal=z.lower();
zFinalList=zFinal.split();

# print(len(zFinalList))


for stopWord in listStopWords:
    for tempStr in zFinalList:
        if(stopWord==tempStr):
            zFinalList.remove(tempStr)

z=tb(' '.join(zFinalList))




files1='WEF_The_future__of_financial_services.pdf'
fileToText1 = convert_pdf_to_txt(files1)
y1=fileToText1
z1=y1.decode("utf-8") 
z1=tb(re.sub('[^A-Za-z]+', ' ', z1))
zFinal=z1.lower();
zFinalList=zFinal.split();

# print(len(zFinalList))


for stopWord in listStopWords:
    for tempStr in zFinalList:
        if(stopWord==tempStr):
            zFinalList.remove(tempStr)

z1=tb(' '.join(zFinalList))
#print(z1)





files2='WEF_The_future_of_financial_infrastructure.pdf'
fileToText2 = convert_pdf_to_txt(files2)
y2=fileToText2
z2=y2.decode("utf-8") 
z2=tb(re.sub('[^A-Za-z]+', ' ', z2))
zFinal=z2.lower();
zFinalList=zFinal.split();

# print(len(zFinalList))


for stopWord in listStopWords:
    for tempStr in zFinalList:
        if(stopWord==tempStr):
            zFinalList.remove(tempStr)

z2=tb(' '.join(zFinalList))
#print(z2)





files3='WEF_A_Blueprint_for_Digital_Identity.pdf'
fileToText3 = convert_pdf_to_txt(files3)
y3=fileToText3
z3=y3.decode("utf-8") 
z3=tb(re.sub('[^A-Za-z]+', ' ', z3))
zFinal=z3.lower();
zFinalList=zFinal.split();

# print(len(zFinalList))


for stopWord in listStopWords:
    for tempStr in zFinalList:
        if(stopWord==tempStr):
            zFinalList.remove(tempStr)

z3=tb(' '.join(zFinalList))
# print(z3)


def tf(word, blob):
    return blob.words.count(word) / len(blob.words)

def n_containing(word, bloblist):
    return sum(1 for blob in bloblist if word in blob)

def idf(word, bloblist):
    return math.log(len(bloblist) / (1 + n_containing(word, bloblist)))

def tfidf(word, blob, bloblist):
    return tf(word, blob) * idf(word, bloblist)

finalList = {}

bloblist = [z, z1, z2, z3]
for i, blob in enumerate(bloblist):
    print("Top words in document {}".format(i + 1))
    scores = {word: tfidf(word, blob, bloblist) for word in blob.words}
    sorted_words = sorted(scores.items(), key=lambda x: x[1])
    for word, score in sorted_words[:40]:
        #print("Word: {}, TF-IDF: {}".format(word, round(score, 5)))
        if word in finalList:
        	print(word)
        finalList.update({word : round(score, 5)})





with open('Top100_WordCountTFIDF.csv', 'w') as f :
	rownum=0
	f.write("%s,%s,%s\n"%("RowId","Word","TF-IDF"))
	for keys,values in finalList.items():
		rownum=rownum+1;
		print(rownum,keys,values)
		f.write("%s,%s,%s\n"%(rownum,keys,values))

