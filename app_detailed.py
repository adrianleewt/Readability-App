"""
Language Analysis - English

This script analyzes .txt files using a number of methods.
"""
#import packages for writing to excel and dataframes
import pandas as pd
from pandas import ExcelWriter, ExcelFile
import xlsxwriter
from io import BytesIO
import os
from assets import *

#change path to appropriate directory
try:
    dirname = os.path.dirname(os.path.abspath(__file__))
    os.chdir(os.path.join(dirname, 'input'))
except:
    print("There is an issue with navigating to the input folder. Please make sure the folder exists in the main directory.")

#get corpus
files = os.listdir()
corpus = get_corpus()

needs_clean = ask_needs_clean()

print("""
================================================================================
Language Analysis to Excel
The script will output an excel file in the output directory with statistics of
your .txt or .docx file. Please note that the filename.xlsx file will be
overwritten if the program is run again. Extract necessary data first before
running the script again. While this script can use both .txt and .docx files,
if possible use .txt to avoid formatting issues creating problems for the
metrics.
================================================================================
""")

#asking user for search terms
search = ask_search()
print("\nThanks! Working... (This may take a bit)\n")
# try:
count = 0
for (data_org,name) in zip(corpus,files):

    #creating excel writer
    filename = name + ".xlsx"
    writer = ExcelWriter(filename)
    count += 1

    #preparing data
    if needs_clean == True:
        data = clean(data_org)
    else:
        data = data_org

    data_word, data_sent = get_tokens(data)

    #basic stat
    stats = stat(data, data_word, data_sent)
    df_stats = pd.DataFrame([stats], columns = ["Syllables", "Words", "Sentences", "Syllables per Word", "Words per Sentence", "Estimated Reading Time (min)", "Flesch Kincaid Grade", "Complex Words", "Wordy Sentences"])
    df_stats.to_excel(writer, "Basic Stats",index = False)

    #common words
    common = most_common(data_word)
    df_common = pd.DataFrame(common, columns = ['Word', 'Frequency'])
    df_common.to_excel(writer, "Common Words", index = False)

    #pos tagging & verbosity
    # w_pos_treebank = pos_treebank(data_word)
    df_verbose = most_verbose(data_word)
    df_verbose.to_excel(writer, "Complex Words", index = False)

    #wordiness
    df_wordy = most_wordy(data_sent)
    df_wordy.to_excel(writer, "Wordy Sentences", index = False)

    #noun Strings
    df_noun_string = noun_string(data_org)
    df_noun_string.to_excel(writer,"Noun Strings (Beta)", index = False)

    #lexdisp
    img_ld = img_lex_disp(data_word,search)

    #fkg_total
    img_fkg = img_fkg_over_text(data_sent)

    #syll_total
    img_syll = img_syll_over_text(data_word)

    #write graphs to excel sheets
    workbook = writer.book
    worksheet = writer.sheets['Basic Stats']

    imgdatas = [img_ld, img_fkg, img_syll]

    for cnt,imgdata in enumerate(imgdatas):

        col = 10 * cnt

        imgdata.seek(0)
        worksheet.insert_image(
            3, col, "",
            {'image_data': imgdata}
        )

    #close and save workbook
    os.chdir(os.path.join(dirname, 'output'))
    writer.save()
    writer.close()

    print("File " + str(count) + " analysis complete.")

print("All done! Please check the main folder for the output files.")

# except:
    # print("Script ran into an error. Please ensure that all filenames are shorter than 251 characters. Please verify the file an directory contents and try again.")
