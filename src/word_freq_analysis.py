import xlrd
import xlwt
import re
import numpy as np
import sys
import codecs
import coding
import copy as cp


def count_words(utterance,dict_words):
    n_words = len(utterance)
    for w in range(n_words):
        # get the key for the word type
        if n_words == 1:
            key = '1syl_ISO'
        elif n_words == 2:
            if w == 0:
                key = '2syl_I'
            elif w == 1:
                key = '2syl_F'
        elif n_words == 3:
            if w == 0:
                key = '3syl_I'
            elif w == 3:
                key = '3syl_F'
            else:
                key = '3syl_M'
        elif n_words > 3:
            if w == 0:
                key = '>3syl_I'
            elif w == n_words:
                key = '>3syl_F'
            else:
                key = '>3syl_M'
        
        # count word frequencies:
        if utterance[w] in dict_words:
            dict_words[utterance[w]][key] += 1
        else:
            dict_words[utterance[w]] = init_dict()
            dict_words[utterance[w]][key] = 1

def write_data(words_dict,table_index,sheet):
    pos_keys = ['1syl_ISO','2syl_I','2syl_F','3syl_I','3syl_M','3syl_F',
             '>3syl_I','>3syl_M','>3syl_F']
    n_cols = len(pos_keys)
    n_words = len(words_dict)
    for i in range(n_cols):
        total = 0
        for j in range(n_words):
            sheet.write(table_index+3+j,i+1,words_dict[words_dict.keys()[j]][pos_keys[i]])
            total += words_dict[words_dict.keys()[j]][pos_keys[i]]
        # word position/type total:
        sheet.write(table_index+3+n_words,i+1,total)
    
    # write row totals
    total = 0
    for i in range(len(words_dict)):
        # write the word
        sheet.write(table_index+3+i,0,words_dict.keys()[i])
        part_sum = sum(words_dict[words_dict.keys()[i]].values()) 
        sheet.write(table_index+3+i,n_cols+1,part_sum)
        total += part_sum
    sheet.write(table_index+3+n_words,n_cols+1,total)
        
    

def init_dict():
    D = {'1syl_ISO':0,'2syl_I':0,'2syl_F':0,'3syl_I':0,'3syl_M':0,'3syl_F':0,
             '>3syl_I':0,'>3syl_M':0,'>3syl_F':0}
    return D

def table_setup(indx,sheet,sesh,dicts):
    ########################
    ### Set up the table ###
    ########################
    sheet.write(indx,0,sesh) # write session at top
    # row titles
    sheet.write_merge(indx+1,indx+2,0,0,'Word')
    sheet.write(indx+3+len(dicts),0,'Total')
    
    #col titles
    sheet.write(indx+1,1,'1 Syl.')
    sheet.write_merge(indx+1,indx+1,2,3,'2 Syl.')
    sheet.write_merge(indx+1,indx+1,4,6,'3 Syl.')
    sheet.write_merge(indx+1,indx+1,7,9,'>3 Syl.')
    sheet.write_merge(indx+1,indx+2,10,10,'Total')
    
    word_locs = ['Isolation','Initial','Final','Initial','Medial','Final','Initial','Medial','Final']
    for i in range(9):
        sheet.write(indx+2,1+i,word_locs[i])
    
def freq_analysis():
    # open coding excel file
    coding_wb = xlrd.open_workbook('../output/Coding_output_utterances.xls')
    coding_sh = coding_wb.sheet_by_index(0)
    # unpack
    participants = coding_sh.col_values(0) 
    sessions = coding_sh.col_values(1)
    orthogs = coding_sh.col_values(2)
    
    wordtoken_wb = xlwt.Workbook()
    n_rows = len(orthogs)
    
    ### gather table data from file
    word_pattern = re.compile("\[[^\[]*\]", re.UNICODE)
    participant = ''
    session = ''
    first = True
    word_dicts = {} # holds a dictionary of word position counts for each word 
    
    for r in range(1,n_rows):
        # if a new participant, start a new sheet
        if participant != participants[r]:
            if not first:
                table_setup(table_index,wordtoken_sh,session,word_dicts)
                write_data(word_dicts,table_index,wordtoken_sh) # fill in the table
            participant = coding_sh.cell_value(r,0) # child string
            session = coding_sh.cell_value(r,1)
            wordtoken_sh = wordtoken_wb.add_sheet(participant)
            table_index = 0
            word_dicts = {}
            first = False
            
        # if starting a new session, write the old one, move to a new table
        elif coding_sh.cell_value(r,1) != session: 
            table_setup(table_index,wordtoken_sh,session,word_dicts)
            write_data(word_dicts,table_index,wordtoken_sh)
            session = coding_sh.cell_value(r,1)
            table_index += 5 + len(word_dicts)
            word_dicts = {}
                    
        utterance_words = re.findall(word_pattern,orthogs[r])
        #if len(utterance_tones_target) == len(utterance_tones_actual) == len( utterance_segments_target) ==  len(utterance_segments_actual):
	count_words(utterance_words, word_dicts)
    
    # write the last table
    table_setup(table_index,wordtoken_sh,session,word_dicts)
    write_data(word_dicts,table_index,wordtoken_sh) # fill in the table
    wordtoken_wb.save('../output/F_WordToken.xls')

def main():
    freq_analysis()
    
if __name__ == "__main__":
    main()
    




