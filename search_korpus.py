#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Fri Oct 13 21:31:23 2017

@author: au183362
"""

# measure how long it takes to run
import time
t0 = time.time()

from collections import OrderedDict
import pandas
import numpy as np
from os.path import join, splitext
from xlutils.copy import copy 
from xlrd import open_workbook 
from splinter import Browser
# initialize (headless) browser
browser = Browser('chrome', headless=True) # NB! Chrome extension must be installed for this to work, otherwise just leave blank, and Firefox will be used as default
#browser = Browser(headless=True)
korpus_url = 'http://ordnet.dk/korpusdk_en/concordance'
# open url
browser.visit(korpus_url)

# use pandas to read the excel file, specifying how many rows to skip before reading the header
Pdir = '/Users/au183362/Dropbox/Parkinson_embodied/stimuli/'
file_name = 'Sammenligningsskema.xls'
rows_to_skip = 17 # for indholdsord, skiprows=62 (NB! doublecheck this)
df = pandas.read_excel(join(Pdir,file_name),skiprows=rows_to_skip) 

# specifying parameters for reading the excel-flie and saving later
search_terms = ('lemma','word')
texts = ('Verb', 'Verb.1','Verb.2', 'Verb.3')
texts_inf = ('Infinitive', 'Infinitive.1','Infinitive.2', 'Infinitive.3')
text_names = ('action', 'non-action', 'act-non-act1', 'act-non-act2')
write_names = (('Automatic','Automatic.1'),('Automatic.2','Automatic.3'),('Automatic.4','Automatic.5'),('Automatic.6','Automatic.7'))
#texts = ('Verb', 'Verb.1')
#texts_inf = ('Infinitive', 'Infinitive.1')
#text_names = ('action', 'non-action')
#write_names = (('Automatic','Automatic.1'),('Automatic.2','Automatic.3'))
word_lists = []
occ = OrderedDict()

# loop over the two old texts (aka. action and non-action)
for num, col in enumerate(texts):
    occ[text_names[num]] = OrderedDict()
    word_list_long = df[col].values
    inf_list_long = df[texts_inf[num]].values
    nans = np.where(pandas.isnull(word_list_long))[0]
    word_lists.append((inf_list_long[:nans[0]-1]))
    word_lists.append((word_list_long[:nans[0]-1]))
    
    for s, term in enumerate(search_terms):  
        html = []
        occ[text_names[num]][term] = OrderedDict() # initializing OrderedDict in order to keep the order of the dictionary as it was defined
        
        for val, words in enumerate(word_lists[s+num*2]):
            words = words.strip()
                
            browser.find_by_name('formal').click()
            browser.find_by_id('search_box').fill('[' + term + '="' + words + '" & pos="V"]')
            browser.find_by_id('search_button').click()
            html.append((browser.html))
            
            reduced = html[val].find('Reduced from ')
            if reduced != -1:
                reduc_end = html[val].find('occurrences',reduced+13)
                occ[text_names[num]][term][words] = int(html[val][reduced+13:reduc_end-1])
            else:
                term_end = html[val].find('occurrences',0)
                term_beg = html[val].rfind('of',0,term_end)
                occ[text_names[num]][term][words] = int(html[val][term_beg+2:term_end])
                
        print(text_names[num])
        print(term)
            
# close (headless) browser
browser.quit()

# save the dictionary as numpy array
np.save('occurrences.npy', occ)

# code needed to load the dictionaries if need be
#occ_lemmas = np.load(join(Pdir,'korpus_dk_search/occ_lemmas.npy')).item()
#occ_words = np.load(join(Pdir,'korpus_dk_search/occ_words.npy')).item()

START_ROW = rows_to_skip+1 # 0 based (subtract 1 from excel row number)
rb = open_workbook(join(Pdir,file_name), formatting_info = True)
r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
wb = copy(rb) # a writable copy (can't read values out of this, only write to it)
w_sheet = wb.get_sheet(0) # the sheet to write to within the writable copy
word_list_index = (1,3)

# iterate over the two texts - first lemmas then words
for num, col in enumerate(write_names):
    for s, term in enumerate(search_terms):
        col_write = np.where(df.columns==write_names[num][s])[0][0] # column number for the relevant "Verb" column
        col_read = np.where(df.columns==texts[num])[0][0] # column number for the relevant "Automatic" column
        for row_index in range(START_ROW, r_sheet.nrows):
            check_word = r_sheet.cell(row_index, col_read).value # get the value of the relevant "Verb" cell to match with those in the word_list
            if (row_index-START_ROW) < len(word_lists[word_list_index[num]]):
                if check_word == word_lists[word_list_index[num]][row_index-START_ROW]:
                    # if check_word and word_lists[num][row_index] are the same, then write the frequency in the relevant cell
                    w_sheet.write(row_index, col_write, occ[text_names[num]][term][word_lists[s+num*2][row_index-START_ROW]])
                    
wb.save(join(Pdir,splitext(file_name)[0]) + '_upd' + splitext(file_name)[-1])


t1 = time.time()
total_time = t1-t0
print(total_time)

