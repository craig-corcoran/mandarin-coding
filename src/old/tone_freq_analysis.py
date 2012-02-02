import xlrd
import xlwt
import re
import numpy as np
import sys
import codecs
import coding
import copy as cp


def count_tones(utterance,d_tone1,d_tone2,d_tone3,d_tone4):
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
        
        # count up tone frequencies
        # Tone 1: level
        if (utterance[w][1] == 'L'):
            d_tone1[key] += 1
        # Tone 2: rising
        elif (utterance[w][1] == 'R'):
            d_tone2[key] += 1
                
        # Tone 3: Fall, Rise
        elif (utterance[w][1:4] == 'FRL') | (utterance[w][1:3] == 'FL'):
            d_tone3[key] += 1
        
        # Tone 4: Fall
        elif (utterance[w][1:3] == 'FH') | (utterance[w][1:3] == 'FM'):
            d_tone4[key] += 1

def write_data(D_tone1,D_tone2,D_tone3,D_tone4,table_index,sheet):
    keys = ['1syl_ISO','2syl_I','2syl_F','3syl_I','3syl_M','3syl_F',
             '>3syl_I','>3syl_M','>3syl_F']
    n_cols = len(keys)
    for i in range(n_cols):
        sheet.write(table_index+3,i+1,D_tone1[keys[i]])
        sheet.write(table_index+4,i+1,D_tone2[keys[i]])
        sheet.write(table_index+5,i+1,D_tone3[keys[i]])
        sheet.write(table_index+6,i+1,D_tone4[keys[i]])
        # word position/type total:
        sheet.write(table_index+7,i+1,D_tone1[keys[i]]+D_tone2[keys[i]]
                                     +D_tone3[keys[i]]+D_tone4[keys[i]]) 
    
    # write row totals
    total = 0
    for i in range(4):
        part_sum = sum(eval("D_tone{0}.values()".format(i+1)))
        sheet.write(table_index+3+i,n_cols+1,part_sum)
        total += part_sum
    sheet.write(table_index+7,n_cols+1,total)
        
    

def init_dicts():
    D_tone1 = {'1syl_ISO':0,'2syl_I':0,'2syl_F':0,'3syl_I':0,'3syl_M':0,'3syl_F':0,
             '>3syl_I':0,'>3syl_M':0,'>3syl_F':0}
    D_tone2 = cp.copy(D_tone1)
    D_tone3 = cp.copy(D_tone1)
    D_tone4 = cp.copy(D_tone1)
    return D_tone1,D_tone2,D_tone3,D_tone4

def table_setup(indx,sheet,sesh):
    ########################
    ### Set up the table ###
    ########################
    sheet.write(indx,0,sesh) # write session at top
    # row titles
    sheet.write_merge(indx+1,indx+2,0,0,'Tone \n Categories')
    for i in range(4):
        sheet.write(indx+3+i,0,'Tone {0}'.format(i+1))
    sheet.write(indx+7,0,'Total')
    
    #col titles
    sheet.write(indx+1,1,'1 Syl.')
    sheet.write_merge(indx+1,indx+1,2,3,'2 Syl.')
    sheet.write_merge(indx+1,indx+1,4,6,'3 Syl.')
    sheet.write_merge(indx+1,indx+1,7,9,'>3 Syl.')
    sheet.write_merge(indx+1,indx+2,10,10,'Total')
    
    word_locs = ['Isolation','Initial','Final','Initial','Medial','Final','Initial','Medial','Final']
    for i in range(9):
        sheet.write(indx+2,1+i,word_locs[i])
    

# open coding excel file
coding_wb = xlrd.open_workbook('../output/coding_output.xls')
coding_sh = coding_wb.sheet_by_index(0)
# unpack
participants = coding_sh.col_values(0) 
sessions = coding_sh.col_values(1)
orthogs = coding_sh.col_values(2)
segment_targets = coding_sh.col_values(3)
segment_actuals = coding_sh.col_values(4)
tone_targets = coding_sh.col_values(5)
tone_actuals = coding_sh.col_values(6)
utt_length = coding_sh.col_values(13) #utterance length

# target tone frequency analysis per session by word position
tonetarget_wb = xlwt.Workbook()
toneactual_wb = xlwt.Workbook()
n_rows = len(tone_targets)

### gather table data from file
word_pattern = re.compile("\[[^\[]*\]", re.UNICODE)
participant = ''
session = ''
first = True
for r in range(1,n_rows):
    # if a new participant, start a new sheet
    if participant != coding_sh.cell_value(r,0):
        if not first:
            try:
                write_data(Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4,table_index,tonetarget_sh) # fill in the table
                write_data(Da_tone1,Da_tone2,Da_tone3,Da_tone4,table_index,toneactual_sh)  
            except:
                print 'error'      
        participant = coding_sh.cell_value(r,0) # child string
        session = coding_sh.cell_value(r,1)
        tonetarget_sh = tonetarget_wb.add_sheet(participant)
        toneactual_sh = toneactual_wb.add_sheet(participant)
        Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4 = init_dicts()
        Da_tone1,Da_tone2,Da_tone3,Da_tone4 = init_dicts()
        table_index = 0
        table_setup(table_index,tonetarget_sh,session)
        table_setup(table_index,toneactual_sh,session)
        first = False
        
    # if starting a new session, write the old one, move to a new table
    elif coding_sh.cell_value(r,1) != session: 
        write_data(Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4,table_index,tonetarget_sh) # fill in the table
        write_data(Da_tone1,Da_tone2,Da_tone3,Da_tone4,table_index,toneactual_sh)
        session = coding_sh.cell_value(r,1)
        Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4 = init_dicts()
        Da_tone1,Da_tone2,Da_tone3,Da_tone4 = init_dicts()
        table_index += 9
        table_setup(table_index,tonetarget_sh,session)
        table_setup(table_index,toneactual_sh,session)
    
    utterance_tones_target = re.findall(word_pattern,tone_targets[r])
    utterance_tones_actual = re.findall(word_pattern,tone_actuals[r])
    utterance_segments_target = re.findall(word_pattern,segment_targets[r])
    utterance_segments_actual = re.findall(word_pattern,segment_actuals[r])
    if len(utterance_tones_target) == len(utterance_tones_actual):
        count_tones(utterance_tones_target,Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4)
        count_tones(utterance_tones_actual,Da_tone1,Da_tone2,Da_tone3,Da_tone4)
    else:
        print 'multiple target tones - not currently processed; row: ' + str(r)

# write the last table
write_data(Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4,table_index,tonetarget_sh) # fill in the table
write_data(Da_tone1,Da_tone2,Da_tone3,Da_tone4,table_index,toneactual_sh)

tonetarget_wb.save('../output/F_ToneTarget.xls')
toneactual_wb.save('../output/F_ToneProduction.xls')
    




