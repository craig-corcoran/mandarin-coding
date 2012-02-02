import xlrd
import xlwt
import re
import numpy as np
import sys
import codecs
import coding
import copy as cp


def update_tone_averages(utterance,accuracy,d_tone1,d_tone2,d_tone3,d_tone4):
    #print 'utterance',utterance
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
            elif w == 2:
                key = '3syl_F'
            else:
                key = '3syl_M'
        elif n_words > 3:
            if w == 0:
                key = '>3syl_I'
            elif w == n_words-1:
                key = '>3syl_F'
            else:
                key = '>3syl_M'
        try:
            # Tone 1: level
            if (utterance[w][1] == 'L'):
                d_tone1[key][0] += float(accuracy[w][1:-1])
                d_tone1[key][1] += 1. # update count
            # Tone 2: rising
            elif (utterance[w][1] == 'R'):
                d_tone2[key][0] += float(accuracy[w][1:-1])
                d_tone2[key][1] += 1. # update count
                    
            # Tone 3: Fall, Rise
            elif (utterance[w][1:4] == 'FRL') | (utterance[w][1:3] == 'FL'):
                d_tone3[key][0] += float(accuracy[w][1:-1])
                d_tone3[key][1] += 1. # update count
            
            # Tone 4: Fall
            elif (utterance[w][1:3] == 'FH') | (utterance[w][1:3] == 'FM'):
                #print 'tone 4 accuracy:, ', float(accuracy[w][1:-1])
                d_tone4[key][0] += float(accuracy[w][1:-1])
                d_tone4[key][1] += 1. # update count
                
        except ValueError as strerror:
            print "Value Error"+str(strerror)
        except:
            print "Unexpected error:", sys.exc_info()[0]

def write_data(D_tone1,D_tone2,D_tone3,D_tone4,table_index,sheet):
    keys = ['1syl_ISO','2syl_I','2syl_F','3syl_I','3syl_M','3syl_F',
             '>3syl_I','>3syl_M','>3syl_F']
    n_cols = len(keys)
    tone1_avg = np.array([0.,0.])
    tone2_avg = np.array([0.,0.])
    tone3_avg = np.array([0.,0.])
    tone4_avg = np.array([0.,0.])
    #print 'table index: ', table_index
    for i in range(n_cols):
        #print 'column: ', i+1
        #print 'tone 4 accuracy: ', '-' if (D_tone1[keys[i]][1]==0) else D_tone1[keys[i]][0]/float(D_tone1[keys[i]][1])
        sheet.write(table_index+3,i+1,'-' if (D_tone1[keys[i]][1]==0) else D_tone1[keys[i]][0]/float(D_tone1[keys[i]][1]))
        sheet.write(table_index+4,i+1,'-' if (D_tone2[keys[i]][1]==0) else D_tone2[keys[i]][0]/float(D_tone2[keys[i]][1]))
        sheet.write(table_index+5,i+1,'-' if (D_tone3[keys[i]][1]==0) else D_tone3[keys[i]][0]/float(D_tone3[keys[i]][1]))
        sheet.write(table_index+6,i+1,'-' if (D_tone4[keys[i]][1]==0) else D_tone4[keys[i]][0]/float(D_tone4[keys[i]][1]))
        tone1_avg += np.array([D_tone1[keys[i]][0],D_tone1[keys[i]][1]])
        tone2_avg += np.array([D_tone2[keys[i]][0],D_tone2[keys[i]][1]])
        tone3_avg += np.array([D_tone3[keys[i]][0],D_tone3[keys[i]][1]])
        tone4_avg += np.array([D_tone4[keys[i]][0],D_tone4[keys[i]][1]])
        # word position/type total:
        sheet.write(table_index+7,i+1,'-' if ((D_tone1[keys[i]][1]+D_tone2[keys[i]][1]
                                     +D_tone3[keys[i]][1]+D_tone4[keys[i]][1])==0)
                                else (D_tone1[keys[i]][0]+D_tone2[keys[i]][0]
        +D_tone3[keys[i]][0]+D_tone4[keys[i]][0])/float(D_tone1[keys[i]][1]+D_tone2[keys[i]][1]
                                     +D_tone3[keys[i]][1]+D_tone4[keys[i]][1]))
    
    total = np.array([0.,0.])
    for i in range(4):
        denom = eval("tone{0}_avg[1]".format(i+1))
        if denom != 0:
            sheet.write(table_index+3+i,n_cols+1,eval("tone{0}_avg[0]/float(tone{0}_avg[1])".format(i+1)))
        else:
            sheet.write(table_index+3+i,n_cols+1,'-')
        total += np.array(eval("[tone{0}_avg[0],tone{0}_avg[1]]".format(i+1)))
    
    if total[1] != 0:
        sheet.write(table_index+7,n_cols+1,total[0]/float(total[1]))
    else:
        sheet.write(table_index+7,n_cols+1,'-')
        
    

def init_dicts():
    #print 'dicts reset'
    D_tone1 = {'1syl_ISO':np.array([0,0.]),'2syl_I':np.array([0.,0.]),'2syl_F':np.array([0.,0.]),'3syl_I':np.array([0.,0.]),'3syl_M':np.array([0.,0.]),'3syl_F':np.array([0.,0.]),
             '>3syl_I':np.array([0,0.]),'>3syl_M':np.array([0.,0.]),'>3syl_F':np.array([0.,0.])}
    D_tone2 = cp.deepcopy(D_tone1)
    D_tone3 = cp.deepcopy(D_tone1)
    D_tone4 = cp.deepcopy(D_tone1)
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
    sheet.write(indx+7,0,'Average')
    
    #col titles
    sheet.write(indx+1,1,'1 Syl.')
    sheet.write_merge(indx+1,indx+1,2,3,'2 Syl.')
    sheet.write_merge(indx+1,indx+1,4,6,'3 Syl.')
    sheet.write_merge(indx+1,indx+1,7,9,'>3 Syl.')
    sheet.write_merge(indx+1,indx+2,10,10,'Weighted \n Average')
    
    word_locs = ['Isolation','Initial','Final','Initial','Medial','Final','Initial','Medial','Final']
    for i in range(9):
        sheet.write(indx+2,1+i,word_locs[i])
    

# open coding excel file
coding_wb = xlrd.open_workbook('../output/coding_output_utterances.xls')
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
utt_length = coding_sh.col_values(13) #utterance length

toneaccuracy_wb = xlwt.Workbook()
segmentaccuracy_wb = xlwt.Workbook()
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
            write_data(Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4,table_index,tone_acc_sh) # fill in the table
            write_data(Ds_tone1,Ds_tone2,Ds_tone3,Ds_tone4,table_index,segm_acc_sh)
        participant = coding_sh.cell_value(r,0) # child string
        session = coding_sh.cell_value(r,1)
        tone_acc_sh = toneaccuracy_wb.add_sheet(participant)
        segm_acc_sh = segmentaccuracy_wb.add_sheet(participant)
        Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4 = init_dicts()
        Ds_tone1,Ds_tone2,Ds_tone3,Ds_tone4 = init_dicts()
        table_index = 0
        table_setup(table_index,tone_acc_sh ,session)
        table_setup(table_index,segm_acc_sh,session)
        first = False
        
    # if starting a new session, write the old one, move to a new table
    elif coding_sh.cell_value(r,1) != session:
        write_data(Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4,table_index,tone_acc_sh) # fill in the table
        write_data(Ds_tone1,Ds_tone2,Ds_tone3,Ds_tone4,table_index,segm_acc_sh)
        session = coding_sh.cell_value(r,1)
        Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4 = init_dicts()
        Ds_tone1,Ds_tone2,Ds_tone3,Ds_tone4 = init_dicts()
        table_index += 9
        table_setup(table_index,tone_acc_sh ,session)
        table_setup(table_index,segm_acc_sh,session)
    
    #print 'row: ',r
    #print 'session: ', session
    
    utterance_tones_target = re.findall(word_pattern,tone_targets[r])
    utterance_segment_accuracy = re.findall(word_pattern,segment_accuracy[r])
    utterance_tone_accuracy = re.findall(word_pattern,tone_accuracy[r])
    if len(utterance_tones_target) == len(utterance_segment_accuracy) == len(utterance_tone_accuracy):
        update_tone_averages(utterance_tones_target,utterance_tone_accuracy,Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4)
        update_tone_averages(utterance_tones_target,utterance_segment_accuracy,Ds_tone1,Ds_tone2,Ds_tone3,Ds_tone4)
    else:
        print 'warning: multiple target tones - not currently processed; row: ' + str(r)

    
# write the last table
write_data(Dt_tone1,Dt_tone2,Dt_tone3,Dt_tone4,table_index,tone_acc_sh) # fill in the table
write_data(Ds_tone1,Ds_tone2,Ds_tone3,Ds_tone4,table_index,segm_acc_sh)

toneaccuracy_wb.save('../output/TA_ToneCategory.xls')
segmentaccuracy_wb.save('../output/SA_ToneCategory.xls')
    




