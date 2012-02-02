import copy
import codecs
import itertools
import numpy as np
import re
import xlrd
import xlwt

# module constants also used by MWCM_coding.py
word_re = '\[[^\[]*\]'
split_re = word_re + '/' + word_re
word_pattern = re.compile(word_re, re.UNICODE)

with codecs.open('../data/vowels',encoding='utf-8') as vowel_file:
    vowels = vowel_file.readline().split(',')

with codecs.open('../data/mandarin_alphabet',encoding='utf-8') as alpha_file:
    IPA_list = alpha_file.readline().split(',')
    
with codecs.open('../data/consonants',encoding='utf-8') as consonant_file:
    consonants = consonant_file.readline().split(',')
   
with codecs.open('../data/velar_consonants',encoding='utf-8') as alpha_file:
    velar_consonants = alpha_file.readline().split(',')

with codecs.open('../data/fricatives',encoding='utf-8') as alpha_file:
    fricatives = alpha_file.readline().split(',')
    
with codecs.open('../data/affricates',encoding='utf-8') as alpha_file:
    affricates = alpha_file.readline().split(',')
    
with codecs.open('../data/aspirated_consonants',encoding='utf-8') as alpha_file:
    aspirated_consonants = alpha_file.readline().split(',')

with codecs.open('../data/liquids',encoding='utf-8') as alpha_file:
    liquids = alpha_file.readline().split(',')

with codecs.open('../data/dipthongs',encoding='utf-8') as alpha_file:
    dipthongs = alpha_file.readline().split(',')

with codecs.open('../data/tripthongs',encoding='utf-8') as alpha_file:
    tripthongs = alpha_file.readline().split(',')


def code_syllable_structure(segments):

    if len(segments) == 0: 
        syllable_structure = 'None'
    else: 
        syllable_structure = ''

    for i in xrange(len(segments)):
        if segments[i] in consonants:
            syllable_structure += 'C'
        elif segments[i] in tripthongs:
            syllable_structure += 'T'
        elif segments[i] in dipthongs:
            syllable_structure += 'D'
        elif segments[i] == vowels[0]: # if a rhotic vowel
            syllable_structure += 'R'
        elif segments[i] in vowels:
            syllable_structure += 'V'
        else:
            print 'Error in syllable structure coding:'
            print 'not a known vowel or consonant'
            print  segments[i]

    return syllable_structure



def code_utterance(utterance_tones_target,utterance_tones_actual,
                  utterance_segments_target,utterance_segments_actual,row, verbose=False):
    n_words = len(utterance_tones_target)
    
    ### CODE WORD POSITION ###
    if n_words == 1:
        word_position = '[ISO]'
    elif n_words > 1:
        word_position = '[I]'+'[M]'*(n_words-2)+'[F]'
    else:
        print 'warning: empty utterance'
        print utterance_segments_target
        
        
    # test tone and segment accuracy and syllable structure for each word in 
    # the utterance
    tone_accuracy = np.array([0.]*n_words)
    segment_accuracy = np.array([0.]*n_words)
    syllable_structure_target = ['']*n_words
    syllable_structure_actual = ['']*n_words
    for w in xrange(n_words):
        
        ### test tone accuracy
        # Tone 1: level
        if (utterance_tones_target[w][1] == 'L'):
            if utterance_tones_actual[w][1] == 'L':
                tone_accuracy[w] = 1.
        # Tone 2: rising
        elif (utterance_tones_target[w][1] == 'R'):
            if utterance_tones_actual[w][1] == 'R':
                tone_accuracy[w] = 1.
                
        # Tone 3: Fall, Rise
        elif (utterance_tones_target[w][1:4] == 'FRL') | (utterance_tones_target[w][1:3] == 'FL'):
            if (utterance_tones_actual[w][1:4] == 'FRL') | (utterance_tones_actual[w][1:3] == 'FL'):
                tone_accuracy[w] = 1.
        elif (utterance_tones_target[w][1:4] == 'FRH') & (utterance_tones_actual[w][1:4] == 'FRH'):
            tone_accuracy[w] = 1.
        elif (utterance_tones_target[w][1:4] == 'FRM') & (utterance_tones_actual[w][1:4] == 'FRM'):
            tone_accuracy[w] = 1.
        
        # Tone 4: Fall
        elif (utterance_tones_target[w][1:3] == 'FH') | (utterance_tones_target[w][1:3] == 'FM'):
            if (utterance_tones_actual[w][1:3] == 'FH') | (utterance_tones_actual[w][1:3] == 'FM'):
                tone_accuracy[w] = 1.
                
        # Tone 5: Neutral 
        elif (utterance_tones_target[w][1] == 'N'):
            pass
            #if utterance_tones_actual[w][1] == 'N':
            #    tone_accuracy[r][w] = 1.
            #else:
            #    tone_accuracy[r][w] = 0.

        # space holding empty bracket []
        elif (utterance_tones_target[w][1] == ']'):
            pass

        else:
            print '*** error - tone not matched'
            print 'row: ', row+1
            print 'tone: ', utterance_tones_target[w]
                
        
         # remove brackets
        target_word = utterance_segments_target[w][1:-1]
        actual_word = utterance_segments_actual[w][1:-1]
        
        # separate segments
        target_segments = find_segments(target_word, IPA_list,row)
        actual_segments = find_segments(actual_word, IPA_list,row)
        num_segments = len(target_segments)
                
        ### Test segment accuracy
        if verbose:
            print 'printing segmentation (verbose on): '
            if num_segments != len(actual_segments):
                print 'target word:', target_word.encode('utf-8')
                print 'actual word:', actual_word.encode('utf-8')
                print 'row, word:', row,' ',w
                print 'target segments:'
                for s in target_segments:
                    print s.encode('utf-8')
                word_re+'/'+word_re
                print 'actual segments:'
                for s in actual_segments:
                    print s.encode('utf-8')
                print actual_segments
                print target_segments
            print 'target segments: '
            for t in target_segments:
                print t.encode('utf-8')
            print 'actual segments: '
            for a in actual_segments:
                print a.encode('utf-8')
            
        # compare segments - take cumulative segment accuracy average
        for i in range(num_segments):
            if i < len(actual_segments): # if actual_segments has the same number of segments
                if target_segments[i] == actual_segments[i]:
                    segment_accuracy[w] += 1

        # normalize the accuracy count
        if num_segments > 0:
            segment_accuracy[w] = segment_accuracy[w]/float(num_segments)

        
        # code the syllable structure

        syllable_structure_target[w] = code_syllable_structure(target_segments)
        syllable_structure_actual[w] = code_syllable_structure(actual_segments)
        
    
    return n_words, word_position, syllable_structure_target, syllable_structure_actual, segment_accuracy, tone_accuracy

def find_segments(word,alpha_list,row):
    '''
     
    find_segments parses a word into IPA phonetic segments according to the valid 
    segments listed in the alpha_list. The segments are appended to the list 'segments' and returned.
    
    ''' 
    lind = 0
    rind = 1
    segments = []
    #print word.encode('utf-8')
    while (lind < len(word)):
        segment = ''
        #print 'word: ',word.encode('utf-8')
        matched = False
         # Note: only allows for segments in the alphabet that have prefixes in the alphabet, with the exception of two characters in between
        while (((word[lind:rind]) in alpha_list)|((word[lind:rind+1]) in alpha_list)|((word[lind:rind+2]) in alpha_list)) & (rind <= len(word)):
            segment = word[lind:rind]
            rind += 1
            matched = True
            
        if rind == lind:
            print 'STUCK'
            print 'current segment: ', segment.encode('utf-8')
            rind += 1
        
        if not matched: # increment and store segment
            print '*** error - segment not matched'
            print 'row: ', row+1
            print 'word: ', word.encode('utf-8')
            segment = word[lind:rind]
            rind += 1
            print 'unknown segment: ',segment.encode('utf-8')
        
        segments.append(segment)
        lind = rind-1
        
    return segments

def find_split_indices(list):
    split_index_list = []
    for i in range(len(list)):
        if '/' in list[i]: 
            split_index_list.append(i)
    return split_index_list

# returns a list of the combinations of options from splits ("/") 
def get_combinations(split_utterance_target,split_indxs):
    option_list = []
    for i in split_indxs:
        option_list.append(tuple(split_utterance_target[i].split('/')))
    combs = []
    for i in itertools.combinations([item for sublist in option_list for item in sublist],len(split_indxs)): 
            combs.append(i)
    return list(set(combs)-set(option_list))

def create_split_options(utterance):
    '''
    from an utterance with splits ('/'), return a list of target utterance options 
    as a list of lists.

    ex: utterance = '[one][two/three]', returns [['[one]','[two]'],['[one]','[three]']]
    '''
    utterance_list = re.findall(re.compile(split_re+'|'+word_re,re.UNICODE),utterance)
    split_indxs = find_split_indices(utterance_list)
    options = []
    if len(split_indxs) > 0:
        # find all permutations of the splits
        split_comb_list = get_combinations(utterance_list,split_indxs)
        for comb in split_comb_list:
            utterance_temp = copy.deepcopy(utterance_list)
            for i in range(len(split_indxs)):
                utterance_temp[split_indxs[i]] = comb[i]
            
            options.append(copy.deepcopy(utterance_temp))
    
        return options

    else:
        print 'create_split_options called on utterance without splits (\'/\')'
        print 'list: ', utterance_list
        return utterance_list


def coding():
    
    wb = xlrd.open_workbook('../data/TXCELA_Chinese_complete112211.xls')
    sh = wb.sheet_by_index(0)
    
    # store all columns for processing or writing to xls file later    
    participants = sh.col_values(0)
    sessions = sh.col_values(1)
    orthogs = sh.col_values(2)
    segment_targets = sh.col_values(3)
    segment_actuals = sh.col_values(4)
    tone_targets = sh.col_values(5)
    tone_actuals = sh.col_values(6)
    notes = sh.col_values(7)
    I = sh.col_values(8)
    J = sh.col_values(9)
    K = sh.col_values(10)
    
    n_rows = len(tone_targets)
    
    # initialize coding rows
    tone_accuracy = [None]*n_rows
    segment_accuracy = [None]*n_rows
    utterance_length = [None]*n_rows
    word_positions = [None]*n_rows
    syllable_structure_target = [None]*n_rows # vowels or consonants
    syllable_structure_actual = [None]*n_rows

    # iterate through each utterance    
    for r in range(1,n_rows): 
        utterance_tones_target = re.findall(word_pattern,tone_targets[r])
        utterance_tones_actual = re.findall(word_pattern,tone_actuals[r])
        utterance_segments_target = re.findall(word_pattern,segment_targets[r])
        utterance_segments_actual = re.findall(word_pattern,segment_actuals[r])
        n_words = len(utterance_tones_actual)
            
        # if there are the same number of words in the basic parse (not accounting for splits)
        if len(utterance_tones_target) == len(utterance_tones_actual) == len(utterance_segments_target) ==  len(utterance_segments_actual):
            # then perform the coding for this row
            utterance_length[r], word_positions[r], syllable_structure_target[r], syllable_structure_actual[r], segment_accuracy[r], tone_accuracy[r] = code_utterance(utterance_tones_target,utterance_tones_actual,
                                                                                                                                                                                            utterance_segments_target,utterance_segments_actual,r)
        else:
            # check for multiple possible tones or segments
            split_utterance_tones_target = re.findall(re.compile(split_re+'|'+word_re,re.UNICODE),tone_targets[r])
            split_utterance_segments_target = re.findall(re.compile(split_re+'|'+word_re,re.UNICODE),segment_targets[r])

            tone_split_indxs = find_split_indices(split_utterance_tones_target)
            if len(tone_split_indxs) > 0:
                tone_comb_list = get_combinations(split_utterance_tones_target,tone_split_indxs)
                for comb in tone_comb_list:
#                    print 'multiple tones'
#                    print 'original tone target: ', tone_targets[r].encode('utf-8')
#                    print 'segmented tone target: ', split_utterance_tones_target
#                    print 'combination: ', comb
#                    print 'row: ', r
#                    print 'indices: ',tone_split_indxs
                    utterance_tones_target_temp = copy.deepcopy(split_utterance_tones_target)
                    for i in range(len(tone_split_indxs)):
                        utterance_tones_target_temp[tone_split_indxs[i]] = comb[i]
                        
                    utterance_length[r], word_positions[r],syllable_structure_target[r], syllable_structure_actual[r], seg_acc, tone_acc = code_utterance(utterance_tones_target_temp,utterance_tones_actual,
                        utterance_segments_target,utterance_segments_actual,r) 
                    if (tone_acc >= np.array(tone_accuracy[r])).all():
                        tone_accuracy[r] = tone_acc
                    if (seg_acc >= np.array(segment_accuracy[r])).all():
                        segment_accuracy[r] = seg_acc
                    
            segment_split_indxs = find_split_indices(split_utterance_segments_target)
            if len(segment_split_indxs) > 0:
                segment_comb_list = get_combinations(split_utterance_segments_target,segment_split_indxs)
                for comb in segment_comb_list:
#                    print 'multiple segments'
#                    print 'original segment target: ', segment_targets[r]
#                    print 'segmented segment target: ', split_utterance_segments_target
#                    print 'combination: ', comb
#                    print 'row: ', r
#                    print 'indices: ',segment_split_indxs
                    utterance_segments_target_temp = copy.deepcopy(split_utterance_segments_target)
                    for i in range(len(segment_split_indxs)):
                        utterance_segments_target_temp[segment_split_indxs[i]] = comb[i]
                    utterance_length[r], word_positions[r],  syllable_structure_target[r], syllable_structure_actual[r], seg_acc, tone_acc = code_utterance(utterance_tones_target,utterance_tones_actual,
                        utterance_segments_target_temp,utterance_segments_actual,r) 
                    if (tone_acc >= np.array(tone_accuracy[r])).all():
                        tone_accuracy[r] = tone_acc
                    if (seg_acc >= np.array(segment_accuracy[r])).all():
                        segment_accuracy[r] = seg_acc
    
    ### export the excel file
    
    export_wb = xlwt.Workbook()
    sheet1 = export_wb.add_sheet("sheet1")

    for r in range(1,n_rows):
        sheet1.write(r,0,participants[r])
        sheet1.write(r,1,sessions[r])
        sheet1.write(r,2,orthogs[r])
        sheet1.write(r,3,segment_targets[r])
        sheet1.write(r,4,segment_actuals[r])
        sheet1.write(r,5,tone_targets[r])
        sheet1.write(r,6,tone_actuals[r])
        sheet1.write(r,7,notes[r])
        sheet1.write(r,8,I[r])
        sheet1.write(r,9,J[r])
        sheet1.write(r,10,K[r])
        sheet1.write(r,11,(str(segment_accuracy[r]).replace(' ','][')).replace('[]',''))
        sheet1.write(r,12,(str(tone_accuracy[r]).replace(' ','][')).replace('[]',''))
        sheet1.write(r,13,str(utterance_length[r]))
        sheet1.write(r,14,str(word_positions[r]))
        sheet1.write(r,15,((str(syllable_structure_target[r]).replace('\'','')).replace(', ','][')).replace('[]',''))
        sheet1.write(r,16,((str(syllable_structure_actual[r]).replace('\'','')).replace(', ','][')).replace('[]',''))
    
    # write titles on first line
    sheet1.write(0,0,participants[0])
    sheet1.write(0,1,sessions[0])
    sheet1.write(0,2,orthogs[0])
    sheet1.write(0,3,segment_targets[0])
    sheet1.write(0,4,segment_actuals[0])
    sheet1.write(0,5,tone_targets[0])
    sheet1.write(0,6,tone_actuals[0])
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
        
    export_wb.save('../output/Coding_output_utterances.xls')
    return export_wb

def main():
    coding()
    
if __name__ == "__main__":
    main()
    
        
