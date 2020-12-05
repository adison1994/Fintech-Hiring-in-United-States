import pdfminer
import re
import xlrd
import pandas as pd
import openpyxl as op
import collections
import operator
import csv
import nltk
nltk.download('punkt') # one time execution
import re
from nltk.stem import WordNetLemmatizer
import string
import numpy as np
import math




from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import BytesIO
from collections import Counter
from nltk.tokenize import word_tokenize

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





files='Beyond_Fintech_-_A_Pragmatic_Assessment_of_Disruptive_Potential_in_Financial_Services.pdf'
fileToText = convert_pdf_to_txt(files)
y=fileToText
z=y.decode("utf-8") 
z=re.sub('[^A-Za-z]+', ' ', z)
#print(z)






files1='WEF_The_future__of_financial_services.pdf'
fileToText1 = convert_pdf_to_txt(files1)
y1=fileToText1
z1=y1.decode("utf-8") 
z1=re.sub('[^A-Za-z]+', ' ', z1)
#print(z1)





files2='WEF_The_future_of_financial_infrastructure.pdf'
fileToText2 = convert_pdf_to_txt(files2)
y2=fileToText2
z2=y2.decode("utf-8") 
z2=re.sub('[^A-Za-z]+', ' ', z2)
#print(z2)





files3='WEF_A_Blueprint_for_Digital_Identity.pdf'
fileToText3 = convert_pdf_to_txt(files3)
y3=fileToText3
z3=y3.decode("utf-8") 
z3=re.sub('[^A-Za-z]+', ' ', z3)
#print(z3)





zFinal=z+z1+z2+z3;
#print(zFinal)






ms = op.load_workbook('Stopwords.xlsx')
ws = ms.active
listStopWords=[]

for row in ws.iter_rows(min_row=1, max_col=1, max_row=1416):
    for cell in row:
        listStopWords.append(cell.value) 

# print(listStopWords)
# print(len(listStopWords))









a=zFinal.lower();
zFinalList=a.split();

# print(len(zFinalList))


for stopWord in listStopWords:
    for tempStr in zFinalList:
        if(stopWord==tempStr):
            zFinalList.remove(tempStr)

# zFinalString=
# for x in zFinalList:
#     zFinalString=zFinalString+" "+x







zFinalString = ' '.join(zFinalList)
#print(zFinalString)

text = word_tokenize(zFinalString)

print("Tokenized Text: \n")
#print(text)



POS_tag = nltk.pos_tag(text)

print("Tokenized Text with POS tags: \n")
#print(POS_tag)


wordnet_lemmatizer = WordNetLemmatizer()

adjective_tags = ['JJ','JJR','JJS']

lemmatized_text = []

for word in POS_tag:
    if word[1] in adjective_tags:
        lemmatized_text.append(str(wordnet_lemmatizer.lemmatize(word[0],pos="a")))
    else:
        lemmatized_text.append(str(wordnet_lemmatizer.lemmatize(word[0]))) #default POS = noun
        
print("Text tokens after lemmatization of adjectives and nouns: \n")
#print lemmatized_text


POS_tag = nltk.pos_tag(lemmatized_text)

print("Lemmatized text with POS tags: \n")
#print(POS_tag)




stopwords = []

wanted_POS = ['NN','NNS','NNP','NNPS','JJ','JJR','JJS','VBG','FW'] 

for word in POS_tag:
    if word[1] not in wanted_POS:
        stopwords.append(word[0])

punctuations = list(str(string.punctuation))

stopwords = stopwords + punctuations

stopword_file = open("long_stopwords.txt", "r")
#Source = https://www.ranks.nl/stopwords

lots_of_stopwords = []

for line in stopword_file.readlines():
    lots_of_stopwords.append(str(line.strip()))

stopwords_plus = []
stopwords_plus = stopwords + lots_of_stopwords
stopwords_plus = set(stopwords_plus)

processed_text = []
for word in lemmatized_text:
    if word not in stopwords_plus:
        processed_text.append(word)
#print(processed_text)


vocabulary = list(set(processed_text))
#print(vocabulary)






vocab_len = len(vocabulary)

weighted_edge = np.zeros((vocab_len,vocab_len),dtype=np.float32)

score = np.zeros((vocab_len),dtype=np.float32)
window_size = 3
covered_coocurrences = []

for i in range(0,vocab_len):
    score[i]=1
    for j in range(0,vocab_len):
        if j==i:
            weighted_edge[i][j]=0
        else:
            for window_start in range(0,(len(processed_text)-window_size)):
                
                window_end = window_start+window_size
                
                window = processed_text[window_start:window_end]
                
                if (vocabulary[i] in window) and (vocabulary[j] in window):
                    
                    index_of_i = window_start + window.index(vocabulary[i])
                    index_of_j = window_start + window.index(vocabulary[j])
                    
                    # index_of_x is the absolute position of the xth term in the window 
                    # (counting from 0) 
                    # in the processed_text
                      
                    if [index_of_i,index_of_j] not in covered_coocurrences:
                        weighted_edge[i][j]+=1/math.fabs(index_of_i-index_of_j)
                        covered_coocurrences.append([index_of_i,index_of_j])





inout = np.zeros((vocab_len),dtype=np.float32)

for i in range(0,vocab_len):
    for j in range(0,vocab_len):
        inout[i]+=weighted_edge[i][j]






MAX_ITERATIONS = 50
d=0.85
threshold = 0.0001 #convergence threshold

for iter in range(0,MAX_ITERATIONS):
    prev_score = np.copy(score)
    
    for i in range(0,vocab_len):
        
        summation = 0
        for j in range(0,vocab_len):
            if weighted_edge[i][j] != 0:
                summation += (weighted_edge[i][j]/inout[j])*score[j]
                
        score[i] = (1-d) + d*(summation)
    
    if np.sum(np.fabs(prev_score-score)) <= threshold: #convergence condition
        print("Converging at iteration "+str(iter)+"....")
        break






for i in range(0,vocab_len):
    print("Score of "+vocabulary[i]+": "+str(score[i]))


phrases = []

phrase = " "
for word in lemmatized_text:
    
    if word in stopwords_plus:
        if phrase!= " ":
            phrases.append(str(phrase).strip().split())
        phrase = " "
    elif word not in stopwords_plus:
        phrase+=str(word)
        phrase+=" "

print("Partitioned Phrases (Candidate Keyphrases): \n")
print(phrases)




unique_phrases = []

for phrase in phrases:
    if phrase not in unique_phrases:
        unique_phrases.append(phrase)

print("Unique Phrases (Candidate Keyphrases): \n")
print(unique_phrases)






for word in vocabulary:
    #print word
    for phrase in unique_phrases:
        if (word in phrase) and ([word] in unique_phrases) and (len(phrase)>1):
            #if len(phrase)>1 then the current phrase is multi-worded.
            #if the word in vocabulary is present in unique_phrases as a single-word-phrase
            # and at the same time present as a word within a multi-worded phrase,
            # then I will remove the single-word-phrase from the list.
            unique_phrases.remove([word])
            
print("Thinned Unique Phrases (Candidate Keyphrases): \n")
print(unique_phrases)

phrase_scores = []
keywords = []
for phrase in unique_phrases:
    phrase_score=0
    keyword = ''
    for word in phrase:
        keyword += str(word)
        keyword += " "
        phrase_score+=score[vocabulary.index(word)]
    phrase_scores.append(phrase_score)
    keywords.append(keyword.strip())

i=0
for keyword in keywords:
    print("Keyword: '"+str(keyword)+"', Score: "+str(phrase_scores[i]))
    i+=1



sorted_index = np.flip(np.argsort(phrase_scores),0)

keywords_num = 10

print("Keywords:\n")

for i in xrange(0,keywords_num):
    print(str(keywords[sorted_index[i]])+", ")