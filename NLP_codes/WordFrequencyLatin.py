import cltk
from cltk.alphabet.lat import LigatureReplacer
import nltk
from cltk.alphabet.lat import JVReplacer
from cltk.alphabet.lat import drop_latin_punctuation
from cltk.lemmatize.backoff import SequentialBackoffLemmatizer
from cltk.lemmatize.processes import LatinLemmatizationProcess
from cltk.lemmatize.lat import LatinBackoffLemmatizer
from cltk.alphabet.lat import remove_accents
from cltk.alphabet.lat import remove_macrons

from cltk.lemmatize.processes import LemmatizationProcess
from cltk.tokenizers import LatinTokenizationProcess
from cltk.languages.utils import get_lang
from cltk.nlp import NLP

from cltk.tokenizers.word import WordTokenizer
from cltk.tokenizers.lat.lat import LatinWordTokenizer
from cltk.tokenizers.lat.lat import LatinLanguageVars
from nltk.tokenize.punkt import PunktLanguageVars

from collections import defaultdict
import pandas as pd
from openpyxl.workbook import Workbook
import numpy as np
import xlsxwriter
import glob


path_files = "Write here the text collection path"

stopwords = []

with open("Latin stopword path",encoding="utf-8") as f:
    stopwords.extend(f.readlines()) # add all the latin  stopwords


with open("Spanish stopword path",encoding="utf-8") as f :
    stopwords.extend(f.readlines())

stopwords = set((sw.strip("\n") for sw in stopwords ))

if __name__ == '__main__':
    pass


def excel_to_dict(path):

    dictionary = {}
    new_dict = {}

    read = pd.read_excel(path)
    pd_headers = read.columns.values.tolist()

    for element in pd_headers:

        clean = read[element].dropna().unique().tolist()
        dictionary[element] = clean

    for k, v in dictionary.items():
        for vi in v:
            new_dict[vi] = k

    return new_dict

correct_lemmas = excel_to_dict("Path to excel file of correct lemmas")
names = excel_to_dict("Path to excel file of Latin and Spanish names")


def normalize(text: str, stopwords: set) -> list:

    replacer = LigatureReplacer()
    tokenizer = LatinWordTokenizer()
    jvreplacer = JVReplacer()

    tok = tokenizer.tokenize(text)
    lower = (word.lower() for word in tok)
    el_sw = (w for w in lower if w not in stopwords)
    replacing = (replacer.replace(word) for word in el_sw)
    replacing2 = (jvreplacer.replace(word) for word in replacing)

    no_punct = []
    for w in replacing2:
        w = drop_latin_punctuation(w)
        w = remove_accents(w)
        w = remove_macrons(w)
        w = w.replace(" ", "").replace("-", "").replace("\n","")
        if len(w) >= 2:
            no_punct.append(w)


    return no_punct


def fixing_lemmas(listLexLem: list,correctList: list ) -> list:

    new_lemmas =[]

    for lex, lemma in listLexLem:
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
    df.index = np.arange(1,len(df)+1) # it starts from index 1
    df = df.rename_axis('ID')

    return df

files = sorted(glob.glob(path_files))


new_LexLemma =[]

for file in files:
    with open(file, encoding="utf-8") as f:

        read = f.read()
        normalized_corpus = normalize(read, stopwords)
        word_lemma = LatinBackoffLemmatizer().lemmatize(normalized_corpus)
        lemmata = [(w,l) for w,l in word_lemma if l!="punc" and w not in stopwords]
        fix1 = fixing_lemmas(lemmata,correct_lemmas) # fixing lemmas
        fix2 = fixing_lemmas(fix1,names) #fixing names
        new_LexLemma.extend(fix2)
        



LemmaLex = [(lemma,lex) for lex,lemma in new_LexLemma if lemma.islower()]                                                                     
cfd = CondFreqDib(LemmaLex,100)
df = cfd_DataFrame(cfd) 

datatoexcel = pd.ExcelWriter('write name of the stylesheet file', engine = "xlsxwriter")
df.to_excel(datatoexcel)
datatoexcel.save()
