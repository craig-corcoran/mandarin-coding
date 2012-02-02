
import xlrd
import xlwt
import re
import numpy as np
import sys
import codecs
import coding
import copy as cp


def update_word_averages(utterance,accuracy,dict_words):
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
        
        # count word frequencies and sum accuracies
        # the dictionaries store [cumulative (undivided) accuracy, total count]
        if utterance[w] in dict_words:
            dict_words[utterance[w]][key][0] += float(accuracy[w][1:-1])
            dict_words[utterance[w]][key][1] += 1. # update count
        else:
            dict_words[utterance[w]] = init_dict()
            dict_words[utterance[w]][key][0] += float(accuracy[w][1:-1])
            dict_words[utterance[w]][key][1] += 1. # update count

def write_data(words_dict,table_index,sheet):
    pos_keys = ['1syl_ISO','2syl_I','2syl_F','3syl_I','3syl_M','3syl_F',
             '>3syl_I','>3syl_M','>3syl_F']
    n_cols = len(pos_keys)
    n_words = len(words_dict)
    
    word_totals = np.array(np.zeros([n_words,2]),dtype=np.float)
    for i in range(n_cols):
        col_avgs = np.array(np.zeros([n_words,2]),dtype=np.float)
        for j in range(n_words):
            sheet.write(table_index+3+j,i+1,'-' if (words_dict[words_dict.keys()[j]][pos_keys[i]][1] == 0) 
                        else words_dict[words_dict.keys()[j]][pos_keys[i]][0]/float(words_dict[words_dict.keys()[j]][pos_keys[i]][1]))
            col_avgs[j,:] += np.array([words_dict[words_dict.keys()[j]][pos_keys[i]][0],words_dict[words_dict.keys()[j]][pos_keys[i]][1]],dtype=np.float)
            word_totals[j,:] += np.array([words_dict[words_dict.keys()[j]][pos_keys[i]][0],words_dict[words_dict.keys()[j]][pos_keys[i]][1]],dtype=np.float)
        # word position/type total:
        sheet.write(table_index+3+n_words,i+1,'-' if (np.sum(col_avgs[:,1])==0)
                                else np.sum(col_avgs[:,0])/float(np.sum(col_avgs[:,1])))
    
    # write row totals
    for i in range(n_words):
        # write the words
        sheet.write(table_index+3+i,0,words_dict.keys()[i])
        sheet.write(table_index+3+i,n_cols+1,'-' if (word_totals[i,1]==0) else word_totals[i,0]/float(word_totals[i,1]))
    sheet.write(table_index+3+n_words,n_cols+1,np.sum(word_totals[:,0])/float(np.sum(word_totals[:,1])))
        
    

def init_dict():
    D = {'1syl_ISO':np.array([0,0.]),'2syl_I':np.array([0.,0.]),'2syl_F':np.array([0.,0.]),'3syl_I':np.array([0.,0.]),'3syl_M':np.array([0.,0.]),'3syl_F':np.array([0.,0.]),
             '>3syl_I':np.array([0,0.]),'>3syl_M':np.array([0.,0.]),'>3syl_F':np.array([0.,0.])}
    return D

def table_setup(indx,sheet,sesh,dicts):
    ########################
    ### Set up the table ###
    ########################
    sheet.write(indx,0,sesh) # write session at top
    # row titles
    sheet.write_merge(indx+1,indx+2,0,0,'Word')
    sheet.write(indx+3+len(dicts),0,'Average')
    
    #col titles
    sheet.write(indx+1,1,'1 Syl.')
    sheet.write_merge(indx+1,indx+1,2,3,'2 Syl.')
    sheet.write_merge(indx+1,indx+1,4,6,'3 Syl.')
    sheet.write_merge(indx+1,indx+1,7,9,'>3 Syl.')
    sheet.write_merge(indx+1,indx+2,10,10,'Weighted \n Average')
    
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
    segment_targets = coding_sh.col_values(3)
    segment_actuals = coding_sh.col_values(4)
    tone_targets = coding_sh.col_values(5)
    tone_actuals = coding_sh.col_values(6)
    segment_accuracy = coding_sh.col_values(11)
    tone_accuracy = coding_sh.col_values(12)
    
    toneaccuracy_wb = xlwt.Workbook()
    segmentaccuracy_wb = xlwt.Workbook()
    
    n_rows = len(orthogs)
    
    ### gather table data from file
    word_pattern = re.compile("\[[^\[]*\]", re.UNICODE)
    participant = ''
    session = ''
    first = True
    tone_dicts = {}
    segment_dicts = {} # holds a dictionary of word position counts for each word 
    
    for r in range(1,n_rows):
        # if a new participant, start a new sheet
        if participant != participants[r]:
            if not first:
                table_setup(table_index,toneaccuracy_sh,session,tone_dicts)
                table_setup(table_index,segmentaccuracy_sh,session,segment_dicts)
                write_data(tone_dicts,table_index,toneaccuracy_sh) # fill in the table
                write_data(segment_dicts,table_index,segmentaccuracy_sh)
            participant = coding_sh.cell_value(r,0) # child string
            session = coding_sh.cell_value(r,1)
            toneaccuracy_sh = toneaccuracy_wb.add_sheet(participant)
            segmentaccuracy_sh = segmentaccuracy_wb.add_sheet(participant)
            table_index = 0
            tone_dicts = {}
            segment_dicts = {}
            first = False
            
        # if starting a new session, write the old one, move to a new table
        elif coding_sh.cell_value(r,1) != session: 
            table_setup(table_index,toneaccuracy_sh,session,tone_dicts)
            table_setup(table_index,segmentaccuracy_sh,session,segment_dicts)
            write_data(tone_dicts,table_index,toneaccuracy_sh) # fill in the table
            write_data(segment_dicts,table_index,segmentaccuracy_sh)
            session = coding_sh.cell_value(r,1)
            table_index += 5 + len(tone_dicts)
            tone_dicts = {}
            segment_dicts = {}
                    
        utterance_words = re.findall(word_pattern,orthogs[r])
        utterance_segment_accuracy = re.findall(word_pattern,segment_accuracy[r])
        utterance_tone_accuracy = re.findall(word_pattern,tone_accuracy[r])
        if len(utterance_words) == len( utterance_segment_accuracy) == len(utterance_tone_accuracy):
        #if True:
            update_word_averages(utterance_words, utterance_tone_accuracy,tone_dicts)
            update_word_averages(utterance_words, utterance_segment_accuracy,segment_dicts)
        else: 
            print 'warning: number of orthographies and segment or tone accuracies differ - not currently processed; row :' + str(r)
    
    # write the last table
    table_setup(table_index,toneaccuracy_sh,session,tone_dicts)
    table_setup(table_index,segmentaccuracy_sh,session,segment_dicts)
    write_data(tone_dicts,table_index,toneaccuracy_sh) # fill in the table
    write_data(segment_dicts,table_index,segmentaccuracy_sh)
    toneaccuracy_wb.save('../output/TA_WordType.xls')
    segmentaccuracy_wb.save('../output/SA_WordType.xls')

def main():
    freq_analysis()
    
if __name__ == "__main__":
    main()
    




