
import os
from assets import clean, cossim
import pandas as pd
from pandas import ExcelWriter, ExcelFile

from assets import get_corpus, clean, cossim, ask_needs_clean

print(
"""
WARNING: This script will output a matrix with the rows and columns in the order
of the input directory. Sort the input to the desired order BEFORE running the
script.
"""
)

#change path to appropriate directory
try:
    dirname = os.path.dirname(os.path.abspath(__file__))
    os.chdir(os.path.join(dirname, 'input'))
except:
    print("There is an issue with navigating to the input folder. Please make sure the folder exists in the main directory.")

#opening files and adding to corpus
corpus_raw = get_corpus()

#clean text
needs_clean = ask_needs_clean()
if needs_clean == True:
    corpus = [clean(text) for text in corpus_raw]
else:
    corpus = corpus_raw

#cossim execute
df_result = cossim(corpus)

#writing and saving xlsx
try:
    os.chdir(os.path.join(dirname, 'output'))
except:
    print("There is an issue with navigating to the output folder. Please make sure the folder exists in the main directory.")

writer = ExcelWriter('CosineSimilarityOutput.xlsx')
df_result.to_excel(writer, "Cosine Similarity",index = True)
writer.save()
writer.close()

print("Done!")
