
#import system and writing functions
import docx2txt
import textstat
import string
import re
import xlsxwriter
from io import BytesIO
import os
from itertools import chain


#import mathy functions
import pandas as pd
from pandas import ExcelWriter, ExcelFile
from collections import Counter
import matplotlib.pyplot as plt
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity
import math

#import nltk and functions
import nltk
nltk.download('punkt')
nltk.download('stopwords')
nltk.download('wordnet')
nltk.download('averaged_perceptron_tagger')
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords
from nltk.corpus import wordnet
from nltk.stem import WordNetLemmatizer


def get_corpus():
    """Returns list of strings from text documents

    - Looks at all files in current working directory.
    - All files in cwd must be .docx or .txt format.
    """
    corpus_raw = []
    files = os.listdir()

    for name in files:
        if ".txt" in name:
            try:
                file = open(name, "rt", encoding='utf8')
                data_org = file.read()
                corpus_raw.append(data_org)
            except:
                print("ERROR: Couldn't open a .txt file. Please ensure that the text is UTF-8 encoded.")
        elif ".docx" in name:
            try:
                data_org = docx2txt.process(name)
                corpus_raw.append(data_org)
            except:
                print("ERROR: Couldn't open a .docx file. Please ensure that the text is UTF-8 encoded.")
        else:
            print("ERROR: Cannot print non .txt or .docx files. Please verify the input folder's contents.")

    return corpus_raw

def clean(data_org):
    """Preprocessing text, returns cleaned text as string.

    - Function needs to be tailored for different types of text.
    - Currently optimized for contracts and policies.
    """
    #cleaning up the text
    data = data_org.lower()
    data = data.replace(';','.')
    data = data.replace(':','.')
    data = data.replace('-',' ')
    data = data.replace('/',' ')
    # data = data.replace('\n',' ')
    data = re.sub(r'\([^)]*\)', '', data)
    pattern = r'\[[^\]]*\]'
    data = re.sub(pattern, '', data)

    if '\n' in data:
        #newline handling
        data = '\r\n'.join([x for x in data.splitlines() if x.strip()])
        data = data.split('\n')
        #removing punctuation at end of line
        data = [x[:-1] for x in data]
        for x in data:
            if x[:-1] in string.punctuation:
                x[:-1]
        #remove digits
        data = [re.sub(r"\d+", "", x) for x in data]
        #remove tabs
        data = [x.replace('\t',' ') for x in data]
        #remove excess spaces
        data = [' '.join(x.split()) for x in data]
        #remove trailing and leading spaces
        data = [x.strip() for x in data]
        #remove empty elements from list
        data = list(filter(None, data))
        #rejoin list into string
        data = '. '.join(data)

    data_list = data.split('. ')
    #remove digits
    data_list = [re.sub(r"\d+", "", x) for x in data_list]
    #strip leading and trailing spaces
    data_list = [x.strip() for x in data_list]
    #remove all extra spaces
    data_list = [' '.join(x.split()) for x in data_list]
    #remove punctuation
    data_list = [x.translate(str.maketrans('', '', string.punctuation)) for x in data_list]
    #filter out none elements
    data_list = list(filter(None, data_list))
    data_list = [x for x in data_list if len(x) > 1]
    data = '. '.join(data_list)

    return data

def get_tokens(data_clean):
    """returns TWO list objects from preprocessed data.

    - data_word is word tokenized. Each element in the list is a word.
    - data_sent is sentence tokenized. Each element in the list is a sentence.
    """
    #sentence tokenization
    data_sent = sent_tokenize(data_clean)
    #tokenizer
    data_tokenized_punc = [word for sent in data_sent for word in nltk.word_tokenize(sent)]
    data_word = [word.lower() for word in data_tokenized_punc if word.isalpha()]

    return data_word, data_sent

def pos_treebank(data_word):
    """Determines part of speech tags, returns list of tuples."""
    #returns dict
    w_pos_treebank = nltk.pos_tag(data_word)
    w_pos_treebank = dict(w_pos_treebank)
    return w_pos_treebank

def get_wordnet_pos(treebank_tag):
    """Converts treebank part of speech tags to wordnet format."""

    if treebank_tag.startswith('J'):
        return wordnet.ADJ
    elif treebank_tag.startswith('V'):
        return wordnet.VERB
    elif treebank_tag.startswith('N'):
        return wordnet.NOUN
    elif treebank_tag.startswith('R'):
        return wordnet.ADV
    else:
        return wordnet.NOUN

def cossim(corpus):
    """Returns dataframe of cosine similarity output."""
    files = os.listdir()
    vectorizer = TfidfVectorizer()
    trsfm = vectorizer.fit_transform(corpus)
    columns = vectorizer.get_feature_names()
    df_tfidf = pd.DataFrame(trsfm.toarray(), columns = columns, index = corpus)
    out = cosine_similarity(trsfm)
    df_result = pd.DataFrame(out, columns = files, index = files)
    return df_result

def ask_search():
    """Returns list of string type search terms."""

    print(
"""
Please enter your desired keywords for the lexical dispersion analysis. For quick templates, enter the following keys:

template_insurance: insurance identifier terms
template_contract: contract identifier terms
template_privacy: privacy contract identifier terms

To stop entering keywords, simply enter an empty input.
"""
        )

    #asking user for search terms
    ask = True
    search = []

    while ask == True:
        temp = input("Enter a keyword: ")
        if temp == "":
            break
        elif temp == "template_insurance":
            search = ["treatment", "premium", "claim", "benefit", "exclusions", "charges", "payment", "occupation"]
            break
        elif temp == "template_contract":
            search = ["defined","liability","service","confidential","terminate","law", "breach"]
            break
        elif temp == "template_privacy":
            search = ["purpose","personal","data","collect","transfer","services","contact","provide","authority","marketing","retention","consent","analysis","analytics"]
            break
        else:
            search.append(temp)

    return search

def ask_needs_clean():
    asking = True

    while asking == True:
        ask = input("Is the data already preprocessed and ready for tokenization? (Y/N) ")
        if ask == "Y":
            needs_clean =  False
            asking = False
        elif ask == "N":
            needs_clean = True
            asking = False
        else:
            print("Invalid Answer. Please enter Y or N")

    return needs_clean

def stat(data, data_word, data_sent):
    """Computes basic overview metrics and returns list of values"""
    #basic counts
    sent = len(data_sent)
    syll = textstat.syllable_count(data)
    word = len(data_word)

    #average calcs
    avg_syll = syll / word
    avg_word = word / sent
    read_time = word/265

    #advance stat
    flesch_kincaid_grade = fkg(int(word), int(sent), int(syll))
    verbose = len([word for word in data_word if textstat.syllable_count(word) > 3])

    wordy = 0
    for item in data_sent:
        token = word_tokenize(item)
        if len(token) > 40:
            wordy += 1
    #writing to list
    stats = [syll,word,sent,avg_syll,avg_word,read_time,flesch_kincaid_grade, verbose, wordy]

    return stats

def fkg(word, sent, syll):
    """flesch kincaid grade calculation. returns float."""
    flesch_kincaid_grade = (0.39* (word / sent)) + (11.8 * (syll / word)) - 15.59
    return flesch_kincaid_grade

def most_common(data_word):
    """Finds most common words and outputs a list of tuples (word, count)."""
    stop_words = set(stopwords.words("english"))

    #filter out stop words
    data_filtered = [word for word in data_word if word not in stop_words]
    cnt = Counter(data_filtered)

    #count most common words
    common = cnt.most_common(100)
    return common

def most_verbose(data_word):
    """Finds long multi-syllable words and outputs a dataframe with values."""

    verbose_words = []
    synonyms = []

    #looping through words to find complex words and their synonyms
    for word in data_word:

        #finding complex words and recording word & lemma
        if textstat.syllable_count(word) > 3:

            word_syn = wordnet.synsets(word)
            lemmas = list(chain.from_iterable([word.lemma_names() for word in word_syn]))
            lemmas = [lemma for lemma in lemmas if textstat.syllable_count(lemma) <= 3]

            verbose_words.append(word)
            synonyms.append(lemmas)


    #creating dataframe with data
    df_verbose = pd.DataFrame({'Word':verbose_words,
                               'Synonyms':synonyms}, columns = ['Word','Synonyms'])

    df_verbose.sort_values('Word',inplace = True)
    df_verbose.drop_duplicates(subset = 'Word', keep = 'first', inplace = True)
    return df_verbose

def most_wordy(data_sent):
    """Finds long sentences and outputs a dataframe with values."""
    #initialize lists
    sylls = []
    words = []
    sents = []
    fkgs = []

    #looping through sentences to find lengthy sentences
    for sent in data_sent:
        token = word_tokenize(sent)
        word = len(token)
        if word > 40:

            #appending to lists
            syll = textstat.syllable_count(sent)
            sylls.append(syll)
            words.append(word)
            sents.append(sent)
            fkgs.append(fkg(int(word), 1, int(syll)))

    #transfer information to dataframe
    df_wordy = pd.DataFrame({'Words' : words,
                          'Syllables' : sylls,
                          'Flesch Kincaid Grade Level': fkgs,
                          'Sentence' : sents}, columns = ["Words", "Syllables", "Flesch Kincaid Grade Level", "Sentence"])
    df_wordy.sort_values("Words", ascending = False, inplace = True)
    return df_wordy

def noun_string(data_org):
    """Finds strings of three nouns or more. Returns DataFrame"""
    chains = []
    tokens = word_tokenize(data_org)
    #tokenize to prepare for tagging
    w_tag = dict(nltk.pos_tag(tokens))
    chain = []
    for w, tag in w_tag.items():
        #find all nouns based on treebank format
        if tag.startswith('N'):
            chain.append(w)
        else:
            if len(chain) >= 3:
                chains.append(" ".join(chain))
            chain = []

    #move information to dataframe for printing to excel
    df_noun_string = pd.DataFrame({'Noun Strings (3+ Nouns in a row)': chains}, columns = ['Noun Strings (3+ Nouns in a row)'])
    return df_noun_string

def lex_disp(data_word,search):
    """Computes x and y values for a lexical dispersion plot. returns two lists."""
    #lexical dispersion
    ld_x = []
    ld_y = []
    for x in range(len(data_word)):
        for y in range(len(search)):
            if data_word[x] == search[y]:
                ld_x.append(x)
                ld_y.append(y)

    df = pd.DataFrame()
    return ld_x, ld_y

def img_lex_disp(data_word,search):
    """Returns BytesIO() object for matplotlib graph based on lex_disp."""

    x_vals, y_vals = lex_disp(data_word,search)

    imgdata = BytesIO()
    fig, ax = plt.subplots()
    ax.plot(x_vals,y_vals,"rx")
    plt.ioff()
    plt.yticks(range(len(search)), search)
    plt.ylim(-1, len(search))
    plt.xlim(-1, len(data_word))
    plt.title("Lexical Dispersion Plot")
    plt.xlabel("Word Offset")
    plt.tight_layout()
    fig.savefig(imgdata, format = 'png', dpi = 200)

    return imgdata

def fkg_over_text(data_sent):
    """Returns two lists of x and y points for an fkg graph"""
    if len(data_sent) >= 200:
        step = 40
    else:
        step = int(len(data_sent)/10)

    y = []
    temp_fkg = []

    for count, sent in enumerate(data_sent, 1):

        temp_fkg.append(sent)

        if count >= step:

            words = [word for sent in temp_fkg for word in nltk.word_tokenize(sent)]
            words = [word.lower() for word in words if word.isalpha()]

            word = len(words)

            syll = sum([textstat.syllable_count(word) for word in words])

            y.append(fkg(word,step,syll))
            temp_fkg = temp_fkg[1:]

    x = range(step,len(y) + step)
    return x,y

def img_fkg_over_text(data_sent):
    """Returns BytesIO() object for matplotlib graph based on fkg_over_text."""
    step = 40

    x,y = fkg_over_text(data_sent)

    imgdata = BytesIO()
    fig, ax = plt.subplots()
    ax.plot(x,y)
    plt.ioff()
    plt.yticks(range(0,int(max(y))+1,int(max(y)/10)))
    plt.ylim(int(min(y)) - 1, int(max(y))+1)
    plt.xlim(0, len(x)+step)
    plt.title("Flesch Kincaid Grade Level (" + str(step) + " Sentence Moving Average)")
    plt.xlabel("Sentence Number")
    plt.tight_layout()
    fig.savefig(imgdata, format = 'png', dpi = 200)

    return imgdata

def syll_over_text(data_word):
    """Returns two lists of x and y points for a syllable per word graph"""

    step = 200
    y = []
    temp_syll = []

    for count, word in enumerate(data_word, 1):

        temp_syll.append(textstat.syllable_count(word))

        if count >= step:
            y.append(sum(temp_syll)/len(temp_syll))
            temp_syll = temp_syll[1:]

    x = range(step,len(y)+step)
    return x,y

def img_syll_over_text(data_word):
    """Returns BytesIO() object for matplotlib graph based on syll_over_text."""

    step = 200

    x,y = syll_over_text(data_word)

    imgdata = BytesIO()
    fig, ax = plt.subplots()
    ax.plot(x,y)
    plt.ioff()
    plt.yticks(range(0,int(max(y) + 2),1))
    plt.ylim(1, 4)
    plt.xlim(step, len(x)+step)
    plt.title("Syllables per Word (" + str(step) + " Word Moving Average)")
    plt.xlabel("Word Number")
    plt.tight_layout()
    fig.savefig(imgdata, format = 'png', dpi = 200)

    return imgdata

def sqrt_round(num):
    """rounds the square root to higher int. returns int. """
    out = math.ceil(math.sqrt(num))
    return out
