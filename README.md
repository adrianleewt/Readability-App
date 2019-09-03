# NLP Readability

These scripts analyze text files using a number of methods. Metrics that are used include simple averages of syllables and words, readability metrics, determining common words, and finding synonyms for complex words. These scripts are compatible with .txt files and .docx files, but .txt files are preferred in order to minimize error in calculating metrics due to different .docx formatting. Please note that unusual formatting in documents (tables, columns, etc) may cause errors and that there is no one size fits all preprocessing method. Therefore, for absolute best performance and accuracy, create a corpus of preprocessed text before running analysis.

###### app_detailed.py
This app produces one excel workbook for each input file. Its output is more in detail and provides specific examples of wordiness to assist in writing. 
###### app_sum.py
This app produces one excel workbook to summarize the analysis of all input files. Its output gives an overview of stats and comparative graphs.
###### app_cossim.py
This app produces a simple cosine similarity matrix to compare the similarity of the input files. 

## Installation

Ensure python is properly installed. If not, see: https://docs.anaconda.com/anaconda/install/

Clone to a local repository. Use the package manager [pipenv](https://docs.pipenv.org/en/latest/install/#installing-pipenv) to install the dependencies. To install these packages in a virtual environment, do the following in the Command Prompt:

```bash
cd Bowtie-Language-Analysis
pipenv install
```

## Usage

With the above packages installed, these scripts can now be run from the command line.

Command Prompt Usage:
1. Place all .txt or .docx files you wish to analyze in the folder named 'input'.
2. Navigate to the parent directory with the cd command.
```
cd Bowtie-Language-Analysis
```
3. Run any of the three apps on the command line with pipenv. You should enter something like:
```
pipenv run python app_sum.py
```
4. The script will ask you if the text is preprocessed or not. If the text is ready for tokenization, enter "Y". If not, enter "N".
5. For app_detailed and app_sum, the script will also ask you for keywords for lexical dispersion analysis. Enter a custom list or enter template keys such as template_insurance or template_contract to access basic lists of keywords.
6. Open the "output" folder in the main directory to access the results.

## Warning
1. Since these scripts rely on writing to excel, ensure that the excel file is not open when attempting to overwrite a result.
2. These scripts currently only support UTF-8. Ensure that text files are in a compatible encoding. 
   This can be done in Save As -> Encoding -> UTF-8
3. Do not change the location of the scripts and folders within the directory without making the reflective changes in the code. This will cause an error.
4. Install with pipenv to avoid updates disrupting dependencies.
