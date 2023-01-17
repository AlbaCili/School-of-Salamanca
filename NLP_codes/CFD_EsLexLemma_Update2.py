from cltk.alphabet.lat import LigatureReplacer
import cltk
import pandas as pd
import csv
from openpyxl.workbook import Workbook
import numpy as np
import xlsxwriter
import glob
import nltk
import spacy
import spacy_spanish_lemmatizer
from spacy.tokenizer import Tokenizer
from spacy.lang.es import Spanish
import string
from collections import defaultdict

path_files = "Write here the path for the text collection"

stopwords = []
names = []

with open("Latin stopword list ",encoding="utf-8") as f:
    stopwords.extend(f.readlines()) # add all the latin  stopwords


with open("Spanish stopword list",encoding="utf-8") as f :
    stopwords.extend(f.readlines())

stopwords = set((sw.strip("\n") for sw in stopwords ))


if __name__=="__main__":
    pass


def excel_to_dict(path):

    dictionary = {}
    new_dict = {}

    read = pd.read_excel(path)
    pd_headers =read.columns.values.tolist()
    #print(pd_headers)
    '''the unique methods avoid that the dropna() drops the entire row
    cleaning the table of nan and putting it into a dictionary'''

    '''the function unique gives an array as output value
   that is why it is better to append the values into a list into a list'''

    for element in pd_headers:

        clean = read[element].dropna().unique().tolist()
        dictionary[element] = clean

    '''we are inverting the keys and the values so 
    that search of the  right lemma is faster and more efficient
    keys = wrong lemmas
    values = right lemmas
    '''
    for k, v in dictionary.items():
        for vi in v:
            new_dict[vi] = k

    return new_dict

correct_lemmas = excel_to_dict("Correct_Lemmas/correctLemmas_es.xlsx")
names = excel_to_dict("StopwordsNames_list/EsLatNames.xlsx")

#print(correct_lemmas)
#print(names)


nlp = Spanish()
tokenizer = Tokenizer(nlp.vocab)
stripping = string.punctuation+"¶"

#this funmction nomralizes the text

def normalize (text: str, stopwords: set ) -> str:

    numbers = [0,1,2,3,4,5,6,7,8,9]
    replacer = LigatureReplacer()
    num_sw =[]
    no_numbers = []

    doc = tokenizer(text)
    tokens =  [token.text for token in doc]
    lower = (word.lower() for word in tokens)

    filter1 = (w for w in lower if w not in stopwords)
    replace = (w.replace("\n","").replace(" ","") for w in filter1)
    no_punct = (w.strip(stripping) for w in replace)
    filter2 =  (replacer.replace(w) for w in no_punct)
    filter3 = [word for word in filter2 if len(word)>=2]


    for num in numbers:
        for w in filter3:
            if str(num) in w:
                num_sw.append(w)

    num_sw = set(num_sw)

    for w in filter3:
        if w not in num_sw:
            no_numbers.append(w)


    return " ".join(no_numbers)


#this function fixes the wrong lemmatization from spacy
def fixing_lemmas(listLexLem: list,correctList: list ) -> list:

    new_lemmas =[]

    for lex, lemma in listLexLem:
        # the function get will check the if lemma is present in the keys
        # if True it will get the values (the right lemma),
        # if False it will get the lemma (second element of the tuple)
        new_lemmas.append((lex, correctList.get(lemma, lemma)))


    return new_lemmas


#calculating the conditional frequency distribution of the tuples (lemma, lexemes)
def CondFreqDib(tuple_list: tuple, most_com: int) -> list:

    dictionary = defaultdict(list)

    cfd_lemmas = nltk.ConditionalFreqDist(tuple_list) # calculating the cfd
    lemmas = sorted(cfd_lemmas.conditions()) # getting the lemmas and sorting
                                             # them in alphabetical order

    for lemma in lemmas:
        sum_values = sum(cfd_lemmas[lemma].values()) # summing up all their frequencies
        keys = [key for key in cfd_lemmas[lemma].keys()]
        dictionary[lemma].extend((keys,sum_values))

    sort_dict = dict(sorted(dictionary.items(),
                            key=lambda lem: lem[1][1],
                            reverse = True))

    mostCommon = [conSamp for conSamp in sort_dict.items()][:most_com]

    return mostCommon


def cfd_DataFrame(list: list) -> pd:

    frequency = [conSamp[1][1] for conSamp in list]
    lexems = [", ".join(conSamp[1][0]) for conSamp in list]
    lemmas = [conSamp[0] for conSamp in list]

    dataFrame = {"Lemma":lemmas,"Frequency":frequency,'Lexems': lexems}
    df = pd.DataFrame(dataFrame) # putting the dictionary into a data frame
    df.index = np.arange(1,len(df)+1) # starting Id 1

    return df


nlp = spacy.load("es_core_news_md")
nlp.max_length = 10000000

files = sorted(glob.glob(path_files))


new_LexLemma =[]

for file in files:

    with open(file, encoding="utf-8") as f:

        read = f.read()
        doc = nlp(normalize(read,stopwords))
        lexlemma = [(token.text, token.lemma_)for token in doc
                    if token.text not in stopwords and token.lemma_ !="punc"]
        fix1 = fixing_lemmas(lexlemma,correct_lemmas) # fixing lemmas
        fix2 = fixing_lemmas(fix1,names) #fixing names
        new_LexLemma.extend(fix2)


LemmaLex = [(lemma,lex) for lex,lemma in new_LexLemma if lemma.islower()] #changing the position of lemmas on index O,                                                                     # position of lexemes on index 1 and filtering the list from names
cfd = CondFreqDib(LemmaLex,100)# calculating the conditional frequency distribution
df = cfd_DataFrame(cfd) #mapping it into a data frame

datatoexcel = pd.ExcelWriter('name of the stylesheet', engine = "xlsxwriter")
df.to_excel(datatoexcel)
datatoexcel.save()



