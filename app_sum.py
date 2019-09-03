
import os
import matplotlib.pyplot as plt
import pandas as pd
from io import BytesIO

from pandas import ExcelWriter, ExcelFile
import openpyxl
from assets import sqrt_round, clean, get_tokens, ask_search, get_corpus,stat,lex_disp, fkg_over_text, syll_over_text, ask_needs_clean

#navigate to input folder and set anchor dirname
try:
    dirname = os.path.dirname(os.path.abspath(__file__))
    os.chdir(os.path.join(dirname, 'input'))
except:
    print("There is an issue with navigating to the input folder. Please make sure the folder exists in the main directory.")

#get corpus
files = os.listdir()
corpus = get_corpus()
needs_clean = ask_needs_clean()

#get search
search = ask_search()
print("\nThanks! Working... (This may take a bit)\n")

dct_stat = {}
width = sqrt_round(len(files))

#initialize imgdatas
imgdata_ld = BytesIO()
imgdata_fkg = BytesIO()
imgdata_syll = BytesIO()

#initialize figures and subplots
fig_ld, ax_ld = plt.subplots(nrows = width, ncols = width, figsize = (20,20))
fig_fkg, ax_fkg = plt.subplots(nrows = width, ncols = width, figsize = (20,20))
fig_syll, ax_syll = plt.subplots(nrows = width, ncols = width, figsize = (20,20))
ind_x = 0
ind_y = 0

#initialize ExcelWriter
filename = 'Data Summary.xlsx'
writer = ExcelWriter(filename)
workbook = writer.book

count = 0
for (data_org,file) in zip(corpus,files):

    if needs_clean == True:
        data = clean(data_org)
    else:
        data = data_org

    data_word, data_sent = get_tokens(data)

    #create stats for data
    stats = stat(data, data_word, data_sent)
    dct_temp = {file:stats}
    dct_stat.update(dct_temp)

    #creating datapoints
    ld_x, ld_y = lex_disp(data_word,search)
    fkg_x, fkg_y = fkg_over_text(data_sent)
    syll_x, syll_y = syll_over_text(data_word)

    #graphing lexical dispersion
    ax_ld[ind_x,ind_y].scatter(ld_x, ld_y)
    ax_ld[ind_x,ind_y].title.set_text(file)
    ax_ld[ind_x,ind_y].set_yticks(range(len(search)))
    ax_ld[ind_x,ind_y].set_ylim(-1,len(search))
    ax_ld[ind_x,ind_y].set_yticklabels(search)

    #graphing fkg plots
    ax_fkg[ind_x,ind_y].plot(fkg_x,fkg_y)
    ax_fkg[ind_x,ind_y].title.set_text(file)
    ax_fkg[ind_x,ind_y].set_yticks(range(0,22,2))
    ax_fkg[ind_x,ind_y].set_ylim(0,22)

    #graphing syllable plots
    ax_syll[ind_x,ind_y].plot(syll_x,syll_y)
    ax_syll[ind_x,ind_y].title.set_text(file)
    ax_syll[ind_x,ind_y].set_yticks(range(0,3))
    ax_syll[ind_x,ind_y].set_ylim(0,3)

    plt.tight_layout()

    ind_y+=1
    if ind_y > width-1:
        ind_y = 0
        ind_x += 1

    count += 1
    print("Stats and Graph Info for file " + str(count) + " collected...")

#write stat dataframe to Excel
out_df = pd.DataFrame.from_dict(dct_stat,'index', columns = ["Syllables", "Words", "Sentences", "Syllables per Word", "Words per Sentence", "Estimated Reading Time (min)", "Flesch Kincaid Grade", "Complex Words", "Wordy Sentences"])
out_df.to_excel(writer, "Basic Stats", index = True)
print("Stats Dataframe complete! Working...")

#save figures
fig_ld.savefig(imgdata_ld, format = 'png', dpi = 150)
fig_fkg.savefig(imgdata_fkg, format = 'png', dpi = 150)
fig_syll.savefig(imgdata_syll, format = 'png', dpi = 150)

sheets = ["Lexical Dispersion Plots", "Flesch Kincaid Grade Plots", "Syllables per Word Plots"]
imgdatas = [imgdata_ld, imgdata_fkg, imgdata_syll]

#write graphs to new sheets in workbook
for sh,imgdata in zip(sheets,imgdatas):
    print("Writing " + sh + " to Excel...")
    worksheet = workbook.add_worksheet(sh)

    imgdata.seek(0)
    worksheet.insert_image(
        0, 0, "",
        {'image_data': imgdata}
    )

#navigate to output, save and close excel file
os.chdir(os.path.join(dirname, 'output'))
writer.save()
writer.close()
print("Done! Check the output folder for the excel file.")
