import nltk
import xlrd

from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize, sent_tokenize

#function for finding tokens in the sentence
def ProperNounExtractor(text):
    sentences = nltk.sent_tokenize(text)
    for sentence in sentences:
        words = nltk.word_tokenize(sentence)
        tagged = nltk.pos_tag(words)

        print("\nAll the question words in the sentence are \n")
        for (word, tag) in tagged:
            if(tag == 'WDT' or tag == 'WP' or tag == 'WP$' or tag == 'WRB' or tag == 'MD'):
                print(word)
        
        print("\nAll the nouns in the sentence are \n")
        for (word, tag) in tagged:
            if(tag == 'NNP' or tag == 'NN' or tag == 'NNS' or tag == 'NNPS'): 
                print(word)
        
        print("\nAll the relations in the sentence are \n")
        for (word, tag) in tagged:
            if(tag == 'VB' or tag == 'VBD' or tag == 'VBG' or tag == 'VBN' or tag == 'VBP' or tag == 'VBZ' or tag == 'EX' or tag== 'JJ' or tag == 'JJR' or tag == 'JJS'):
                print(word)

#function for reading the excel file
loc = ".//form.xlsx"
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0,0)

for i in range(sheet.nrows):
        text = sheet.cell_value(i,1)
        print("\n------------------------------ {} ------------------------------\n".format(i))
        print(sheet.cell_value(i,1))
        ProperNounExtractor(text)