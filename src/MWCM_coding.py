import copy
import codecs
import itertools
import numpy as np
import re
import xlrd
import xlwt
import coding
import sys

def measure_word_complexity(syl_struct,word,tone):
    ''' Perform word complexity measure for syllable structure, word segments, and tone '''

    segments = coding.find_segments(word[1:-1], coding.IPA_list,0)
    WCM_S = [None]*10
    WCM_S[0] = 0 # first parameter always zero (for now)
     
    # param 1 : one point if it ends in a consonant (vowel, consonant coding already done in utterance coding
    if syl_struct[-2] == "C":
        WCM_S[1] = 1
    else:
        WCM_S[1] = 0
    
    # add a point for each velar consonant
    WCM_S[2] = 0
    for vc in coding.velar_consonants:
        WCM_S[2] += segments.count(vc)
    
    # add a point for each fricative    
    WCM_S[3] = 0
    for f in coding.fricatives:        
#        print 'segments ', segments
#        print f
#        print 'count ',segments.count(f)
        WCM_S[3] += segments.count(f)
    
    # two points for each affricate 
    WCM_S[4] = 0
    for af in coding.affricates:
        WCM_S[4] += 2*segments.count(af)
#        print 'segments ', segments
#        print af
#        print 'count ',segments.count(af)
        
    WCM_S[5] = 0
    for ac in coding.aspirated_consonants:
        WCM_S[5] += segments.count(ac)
        
    WCM_S[6] = 0 
    for lc in coding.liquids:
        WCM_S[6] += segments.count(lc)
        
    WCM_S[7] = 0
    for dt in coding.dipthongs:
        WCM_S[7] += segments.count(dt)
        
    WCM_S[8] = 0
    for tt in coding.tripthongs:
        WCM_S[8] += 2*segments.count(tt)

    # add non-standard dipthong and tripthong
    if syl_struct.count('VVV') > 0:
        WCM_S[8] += 2*syl_struct.count('VVV')
    elif syl_struct.count('VV') > 0:
        WCM_S[7] += syl_struct.count('VV')

    rv = u'\u025a' # coding.vowels[0] # rhotic vowel should be first in vowel list
    WCM_S[9] = segments.count(rv)
    
    # total the segment score
    WCM_S_total = sum(WCM_S) 
                   
    # Tone 2 - 1pt
    if (tone[1] == 'R'):
        WCM_T = 1
    # Tone 3 - 2pts
    elif (tone[1:3] == 'FR') | (tone[1:3] == 'FL'): 
        WCM_T = 2
    else:
        WCM_T = 0  
        
    return WCM_S, WCM_S_total, WCM_T


def wcm_single(WCM_S, WCM_S_total, WCM_T, syl_struct,segments,tone):
    '''appends the word complexity measure for syl_struct, segments, and tone to the WCM lists '''
    wcm_s, wcm_s_total, wcm_t = measure_word_complexity(syl_struct, segments, tone)

    for i in xrange(len(WCM_S)):
        WCM_S[i].append(wcm_s[i])
    
    WCM_S_total.append(wcm_s_total)
    WCM_T.append(wcm_t)

def wcm_multi(WCM_S, WCM_S_total, WCM_T, syl_struct, tone, segment_options, w):
    ''' Chooses between mutliple target segment options for the max total complexity
    measure and appends to WCM lists. Here it is assumed that the syllable structure
    does not change between options, and that there are never multiple tone options'''

    max_total = -1
    for seg_op in segment_options:
        _wcm_s, _wcm_s_total, _wcm_t = measure_word_complexity(syl_struct, seg_op[w], tone)
        if (_wcm_s_total > max_total):
            wcm_s = copy.copy(_wcm_s)
            wcm_s_total = copy.copy(_wcm_s_total)
            wcm_t = copy.copy(_wcm_t)
            max_total = wcm_s_total

    for i in xrange(len(WCM_S)):
        WCM_S[i].append(wcm_s[i])
    
    WCM_S_total.append(wcm_s_total)
    WCM_T.append(wcm_t)
    

def MWCM_coding():
    
    # start from coding output, going to copy most and split into one line per word (instead of utterance)
    wb = xlrd.open_workbook('../output/Coding_output_utterances.xls') 
    sh = wb.sheet_by_index(0)
    
    # store all columns for processing or writing to xls file later    
    participants = sh.col_values(0)
    sessions = sh.col_values(1)
    orthogs = sh.col_values(2)
    segments_target = sh.col_values(3)
    segments_actual = sh.col_values(4)
    tones_target = sh.col_values(5)
    tones_actual = sh.col_values(6)
    notes = sh.col_values(7)
    I = sh.col_values(8)
    J = sh.col_values(9)
    K = sh.col_values(10)
    segment_accuracy = sh.col_values(11)
    tone_accuracy = sh.col_values(12)
    length = sh.col_values(13)
    word_position = sh.col_values(14)
    syllable_structure_target = sh.col_values(15)
    syllable_structure_actual = sh.col_values(16)
    
    n_rows = len(tones_target)
    
    # initialize output lists, which will be the length of the total number of words
    # rather than the number of utterances
    out_participants = []
    out_sessions = []
    out_orthogs = []
    out_segments_target = []
    out_segments_actual = []
    out_tone_target = []
    out_tone_actual = [] 
    out_notes = []
    out_I = []
    out_J = []
    out_K = []
    out_segment_accuracy = []
    out_tone_accuracy = []
    out_length = []
    out_position = []
    out_syllable_struct_target = []
    out_syllable_struct_actual = []
    WCM_S_target = [[],[],[],[],[],[],[],[],[],[]]
    WCM_S_actual = [[],[],[],[],[],[],[],[],[],[]]
    WCM_S_target_total = []
    WCM_S_actual_total = []
    WCM_T_target = []
    WCM_T_actual = []
    
    word_pattern = coding.word_pattern
    # iterate through each utterance    
    for r in range(1,n_rows):
        utterance_orthogs = re.findall(word_pattern, orthogs[r])
        utterance_segments_target = re.findall(word_pattern,segments_target[r])
        utterance_segments_actual = re.findall(word_pattern,segments_actual[r])
        utterance_tones_target = re.findall(word_pattern,tones_target[r])
        utterance_tones_actual = re.findall(word_pattern,tones_actual[r])
        utterance_segment_accuracy = re.findall(word_pattern,segment_accuracy[r])
        utterance_tone_accuracy = re.findall(word_pattern,tone_accuracy[r])
        utterance_position = re.findall(word_pattern,word_position[r])
        utterance_syllable_struct_target = re.findall(word_pattern, syllable_structure_target[r])
        utterance_syllable_struct_actual = re.findall(word_pattern, syllable_structure_actual[r])
        n_words = len(utterance_tones_actual)
        
    
        if not (len(utterance_tones_target) == len(utterance_tones_actual) == \
                len(utterance_segments_target) ==  len(utterance_segments_actual)):

            segment_target_options = coding.create_split_options(segments_target[r])
            split_indxs,split_segtarg_utter_list = coding.find_split_indices(segments_target[r]) # find splits to write later
        
        for w in xrange(n_words):
            
            try:
            
                # copy the row for each word for columns that aren't split into words
                out_participants.append(participants[r])
                out_sessions.append(sessions[r])
                out_notes.append(notes[r])
                out_I.append(I[r])
                out_J.append(J[r])
                out_K.append(K[r])
                out_length.append(length[r])
                
                # split the words into separate elements for each row
                # print 'row: ', r+1, ' word: ', w+1
                out_orthogs.append(utterance_orthogs[w])
               
                out_segments_actual.append(utterance_segments_actual[w])
                out_tone_target.append(utterance_tones_target[w])
                out_tone_actual.append(utterance_tones_actual[w])
                out_segment_accuracy.append(utterance_segment_accuracy[w])
                out_tone_accuracy.append(utterance_tone_accuracy[w])
                out_position.append(utterance_position[w])
                out_syllable_struct_target.append(utterance_syllable_struct_target[w])
                out_syllable_struct_actual.append(utterance_syllable_struct_actual[w])

                wcm_single(WCM_S_actual,WCM_S_actual_total,WCM_T_actual,utterance_syllable_struct_actual[w], 
                                                    utterance_segments_actual[w], utterance_tones_actual[w])

                if len(utterance_tones_target) == len(utterance_tones_actual) == \
                   len(utterance_segments_target) ==  len(utterance_segments_actual):
                    
                    out_segments_target.append(utterance_segments_target[w])
                    wcm_single(WCM_S_target,WCM_S_target_total,WCM_T_target,utterance_syllable_struct_target[w],
                                                        utterance_segments_target[w], utterance_tones_target[w])

                else:
                    # segment_target_options defined above
                    # take the max of the word complexities for the target
                    out_segments_target.append(split_segtarg_utter_list[w])

                    wcm_multi(WCM_S_target,WCM_S_target_total,WCM_T_target,utterance_syllable_struct_target[w],
                                                        utterance_tones_target[w], segment_target_options, w)

            except Exception as e: 
                print 'Error in row ',r+1,' word ', w+1
                print e.args
                pass


    ### export the excel file
    export_wb = xlwt.Workbook()
    
    # The workbooks is empty, so you have to add a sheet *
    sheet1 = export_wb.add_sheet("word coding")
    n_words_uttered = len(out_orthogs)

    for w in xrange(0, n_words_uttered): # TODO should be n_words +1?
        sheet1.write(w+1,0,out_participants[w])
        sheet1.write(w+1,1,out_sessions[w])
        sheet1.write(w+1,2,out_orthogs[w])
        sheet1.write(w+1,3,out_segments_target[w])
        sheet1.write(w+1,4,out_segments_actual[w])
        sheet1.write(w+1,5,out_tone_target[w])
        sheet1.write(w+1,6,out_tone_actual[w])
        sheet1.write(w+1,7,out_notes[w])
        sheet1.write(w+1,8,out_I[w])
        sheet1.write(w+1,9,out_J[w])
        sheet1.write(w+1,10,out_K[w])
        sheet1.write(w+1,11,(str(out_segment_accuracy[w]).replace(' ','][')).replace('[]','')) # TODO need?
        sheet1.write(w+1,12,(str(out_tone_accuracy[w]).replace(' ','][')).replace('[]',''))
        sheet1.write(w+1,13,str(out_length[w]))
        sheet1.write(w+1,14,str(out_position[w]))
        sheet1.write(w+1,15,((str(out_syllable_struct_target[w]).replace('\'','')).replace(', ','][')).replace('[]',''))
        sheet1.write(w+1,16,((str(out_syllable_struct_actual[w]).replace('\'','')).replace(', ','][')).replace('[]',''))
        
        # for the 10 params, write WCM_S_target
        for i in xrange(10):
            sheet1.write(w+1,17+i,str(WCM_S_target[i][w]))
        sheet1.write(w+1,27,str(WCM_S_target_total[w]))
        
        for i in xrange(10):
            sheet1.write(w+1,28+i,str(WCM_S_actual[i][w]))
        sheet1.write(w+1,38,str(WCM_S_actual_total[w]))
        
        sheet1.write(w+1,39,str(WCM_T_target[w]))
        sheet1.write(w+1,40,str(WCM_T_actual[w]))  
    
    # write titles on first line
    sheet1.write(0,0,participants[0])
    sheet1.write(0,1,sessions[0])
    sheet1.write(0,2,orthogs[0])
    sheet1.write(0,3,segments_target[0])
    sheet1.write(0,4,segments_actual[0])
    sheet1.write(0,5,tones_target[0])
    sheet1.write(0,6,tones_actual[0])
    sheet1.write(0,7,notes[0])
    sheet1.write(0,8,I[0])
    sheet1.write(0,9,J[0])
    sheet1.write(0,10,K[0])
    sheet1.write(0,11,"Segment Accuracy")
    sheet1.write(0,12,"Tone Accuracy")
    sheet1.write(0,13,"Length")
    sheet1.write(0,14,"Position")
    sheet1.write(0,15,"Syllable Structure-Target")
    sheet1.write(0,16,"Syllable Structure-Actual")
    
    # for the 10 params, write WCM_S_target
    for i in xrange(27-17):
        sheet1.write(0,17+i,'MWCM_Starget '+ str(i+1))
    #sheet1.write_merge(0,0,17,26,'MWCM_Starget')
    sheet1.write(0,27,'Total MWCM_Starget')
    for i in xrange(38-28):
        sheet1.write(0,28+i,'MWCM_Sactual '+ str(i+1))
    #sheet1.write_merge(0,0,28,37,'MWCM_Sactual')
    sheet1.write(0,38,'Total MWCM_Sactual')
    sheet1.write(0,39,'MWCM_Ttarget')
    sheet1.write(0,40,'MWCM_Tactual') 
        
    export_wb.save('../output/Coding_output_words.xls')

def test_rhotic():
    wb = xlrd.open_workbook('../data/rhotic_test.xls') 
    sh = wb.sheet_by_index(0)
    
    # store all columns for processing or writing to xls file later    
    segments_target = sh.col_values(3)
    segments_actual = sh.col_values(4)
    n_rows = len(segments_actual)
    
    rhotic_vowel = u'\u025a' #coding.vowels[0] 
    print 'rhotic vowel: ', rhotic_vowel
    for i in xrange(n_rows):
        assert rhotic_vowel in segments_target[i]
        assert not (rhotic_vowel in segments_actual[i])
def main():
    MWCM_coding()
    
if __name__ == "__main__":
    main()
