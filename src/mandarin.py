import xlrd
import xlwt
import sys
import copy

pos_keys = ['1[ISO]','2[I]','2[F]','3[I]','3[M]','3[F]',
             '4[I]','4[M]','4[F]']

class Word(object):

    '''takes a ColumnDictionary object and stores the word tone, segment and 
    tone accuracies, and MWCM values. The Column dictionary should be built
    using the Coding_output_words'''

    def __init__(self, col_dict, row, tone="target"):
        
        self.participant = col_dict["Participant"]
        self.session = col_dict["Session"]
            
        # get tone and tone number
        self.tone_target = col_dict["Tone Target"][row]
        self.tone_actual = col_dict["Tone Actual"][row]

        if tone == "target":
            tone_list = self.tone_target
        else:
            tone_list = self.tone_actual

        # set tone number according to tone variable
        # Tone 1: level
        if (tone_list[1] == 'L'):
            self.tone_number = 1
        # Tone 2: rising
        elif (tone_list[1] == 'R'):
            self.tone_number = 2
        # Tone 3: Fall, Rise
        elif (tone_list[1:4] == 'FRL') | (tone_list[1:3] == 'FL'):
            self.tone_number = 3
        # Tone 4: Fall
        elif (tone_list[1:3] == 'FH') | (tone_list[1:3] == 'FM'):
            self.tone_number = 4 
        else:
            self.tone_number = 5
         
        # everything else:
        self.position = col_dict["Position"][row]
        self.orthography = col_dict["Orthography"][row]
        # convert below to numeric types?
        self.length = int(col_dict["Length"][row])
        self.segment_accuracy = col_dict["Segment Accuracy"][row]
        self.tone_accuracy = col_dict["Tone Accuracy"][row]

        self.MWCM_Starget = col_dict["Total MWCM_Starget"][row]
        self.MWCM_Sactual = col_dict["Total MWCM_Sactual"][row]
        self.MWCM_Ttarget = col_dict["MWCM_Ttarget"][row]
        self.MWCM_Tactual = col_dict["MWCM_Tactual"][row]

    def __hash__(self):
        return hash(self.orthography)

    def __equal__(self, other):
        return self.orthography == other.orthography


class ColumnDictionary(object):
    '''takes the path to a workbook string and creates a dictionary for the 
    column names as keys'''

    def __init__(self, workbook_string):
        
        wb = xlrd.open_workbook(workbook_string)
        self.sh = wb.sheet_by_index(0)
        self.keys = self.sh.row_values(0)

    def __getitem__(self, column_name):
        
        col_ind = self.keys.index(column_name)
        return self.sh.col_values(col_ind)

class Table(object):
    ''' consists of two dictionaries for sorting by tone and by word. '''
    def __init__(self, session):
        #print 'new table'

        self.tone_dict = {}
        for i in xrange(4):
            self.tone_dict[i+1] = dict(zip(pos_keys,[(0,0)]*len(pos_keys))) 

        self.word_dict = {}
        self.session = session
        

    def add_word(self,word, value):
        ''' update the table with the given word '''
        #print 'adding a word to a table'
            
        # initialize position dictionary
        if not self.word_dict.has_key(word.orthography):
            self.word_dict[word.orthography] = dict(zip(pos_keys,[(0,0)]*len(pos_keys))) 
    
        if not self.tone_dict.has_key(word.tone_number):
            self.tone_dict[word.tone_number] = dict(zip(pos_keys,[(0,0)]*len(pos_keys)))
    
        length = word.length # length of the utterance the word appears in
        if int(length) > 4:
            length = 4
        pos_key = str(length)+str(word.position)

        if value < 1e-5: value = 0
            
        if not self.word_dict[word.orthography].has_key(pos_key):
            self.word_dict[word.orthography][pos_key] = (float(value),1)
        else:
            pair =  copy.deepcopy(self.word_dict[word.orthography][pos_key])
            val = float(pair[0])+float(value)
            if val  < 1e-5: val = 0
            self.word_dict[word.orthography][pos_key] = (val, int(pair[1])+1)

        if not self.tone_dict[word.tone_number].has_key(pos_key):
            self.tone_dict[word.tone_number][pos_key] = (float(value),1)
         
        else:
            pair = copy.deepcopy(self.tone_dict[word.tone_number][pos_key])
            val = float(pair[0])+float(value)
            if val  < 1e-5: val = 0
            self.tone_dict[word.tone_number][pos_key] = (val, int(pair[1])+1)

        # don't count the values from words with unknown tone, but keep in table
        if word.tone_number > 4:
            self.tone_dict[word.tone_number][pos_key] = (0,0)
            self.word_dict[word.orthography][pos_key] = (0,0)
        

def write_table(table, sheet, tl_index, averaging = True, sorting = "tone", col_dict = None):
    num_tones = 4 #len(table.tone_dict.keys())
    # total number of words across all tones
    
    word_list = table.word_dict.keys()

    #print 'word list: ', word_list
    num_words = len(word_list)
    if sorting == "word":
        num_rows = num_words
    elif sorting == "tone":
        num_rows = num_tones
    ########################
    ### Set up the table ###
    ########################
    sheet.write(tl_index,0,table.session) # write session at top
    # row titles
    sheet.write_merge(tl_index+1,tl_index+2,0,0,'Word')
    if averaging:
        sheet.write(tl_index + num_rows + 3 ,0,'Average')
    else:
        sheet.write(tl_index + num_rows + 3 ,0,'Total')
    
    #print 'num words: ', num_words
    #print 'num tones: ', num_tones

    #col titles
    sheet.write(tl_index+1,1,'1 Syl.')
    sheet.write_merge(tl_index+1,tl_index+1,2,3,'2 Syl.')
    sheet.write_merge(tl_index+1,tl_index+1,4,6,'3 Syl.')
    sheet.write_merge(tl_index+1,tl_index+1,7,9,'>3 Syl.')
    if averaging:
        sheet.write_merge(tl_index+1,tl_index+2,10,10,'Weighted Average')
    else:
        sheet.write_merge(tl_index+1,tl_index+2,10,10,'Total')

    word_locs = ['Isolation','Initial','Final','Initial','Medial','Final','Initial','Medial','Final']
    for i in range(9):
        sheet.write(tl_index+2,1+i,word_locs[i])

    ########################
    ### Write Table Data ###
    ########################

    #pos_inds = zip(pos_keys, range(len(pos_keys)))
    col_count = dict(zip(pos_keys,[0]*len(pos_keys)))
    col_total = dict(zip(pos_keys,[0]*len(pos_keys)))


    for t in xrange(num_rows): # for all words/tones
        
        word = word_list[t]

        if sorting == "tone":
            if t < 4:
                sheet.write(tl_index + 3 + t, 0, "Tone "+str(t+1))
            else:
                pass #sheet.write(tl_index + 3 + t, 0, "Other")

        elif sorting == "word":
            sheet.write(tl_index + 3 + t, 0, word_list[t])

        row_count = 0
        row_total = 0

        for n in xrange(len(pos_keys)):

            pk = pos_keys[n]
            if sorting == "word":
                if table.word_dict.has_key(word):
                    out = table.word_dict[word].get(pk)
                    if out is None: out = (0,0)
                else:
                    out = (0,0)

            elif sorting == "tone":
                if table.tone_dict.has_key(t+1):
                    out = table.tone_dict[t+1].get(pk)
                    if out is None: out = (0,0)
                else:
                    out = (0,0)

            value,count = out
            #print value, count
            value = float(value)
            count = int(count)
                
            row_count += count
            row_total += value # accumulate word total
            col_count[pk] += count
            col_total[pk] += value #  accumulated column (pos) total

            # write the value of this word in this position
            if averaging & (count != 0):
                value = value/count
            sheet.write(tl_index + t + 3, n+1, value) 

        # write the row total for the word
        if averaging & (row_count != 0):
            row_total = row_total/row_count
        sheet.write(tl_index + t + 3, len(pos_keys)+1, row_total)
            
    # write the tone total for all positions
    for i in xrange(len(pos_keys)):
        pk = pos_keys[i]
        if averaging & (col_count[pk] != 0):
            col_total[pk] = col_total[pk]/col_count[pk]
        sheet.write(tl_index + 3 + num_rows, i+1, col_total[pk])

    # grand total for session
    total = sum(col_total.values())
    if averaging:
        total = total/len(col_total.values())

    sheet.write(tl_index + 3 + num_rows, len(pos_keys)+1, total)
    
    return tl_index + num_rows + 5 


def FrequencyAnalysis(use_tone= "target", table_sorting = "tone"):
    # to be analysis code
    col_dict = ColumnDictionary('../output/Coding_output_words.xls')
    participant_list = col_dict["Participant"][1:]
    session_list = col_dict["Session"][1:]
    curr_sesh = None
    curr_table = None
    curr_participant = None
    wb = xlwt.Workbook()
    tl_index = 0
    sheet = None

    # for all words in all sessions
    for s in xrange(len(session_list)):

        # if starting a new session
        if session_list[s] is not curr_sesh:

            if s is not 0: # if not first line
                tl_index = write_table(curr_table, sheet, tl_index, averaging = False, sorting = table_sorting) # write the old table
            
            # if it is a new participant, make a new sheet
            if participant_list[s] is not curr_participant:
                print 'new sheet: ', participant_list[s]
                curr_participant = participant_list[s]
                sheet = wb.add_sheet(curr_participant)
                tl_index = 0

            # make new table for new session
            curr_sesh = session_list[s]
            curr_table = Table(curr_sesh)

        curr_table.add_word( Word(col_dict,s+1,tone=use_tone), 1) # increase freq count
        
        
    # write the last table
    write_table(curr_table, sheet, tl_index, averaging = False, sorting = table_sorting) 
        
    # write the excel 
    if table_sorting == "tone":
        if use_tone == "target":
            wb.save('../output/F_ToneTarget.xls')
        elif (use_tone == "actual") | (use_tone == "production"):
            wb.save('../output/F_ToneProduction.xls')

    if table_sorting == "word":
            wb.save('../output/F_WordToken.xls')


def MWCM(use_tone= "target", measuring="target", table_sorting = "tone"):
    
    
    # to be analysis code
    #col_dict = ColumnDictionary('../output/Coding_output_words_test.xls')
    col_dict = ColumnDictionary('../output/Coding_output_words.xls')
    participant_list = col_dict["Participant"][1:]
    session_list = col_dict["Session"][1:]
    curr_sesh = None
    curr_table = None
    curr_participant = None
    sheet_total=sheet_average=None
    wb = xlwt.Workbook()
    tl_index = 0
    tl_index_avg = 0

    # for all words in all sessions
    for s in xrange(len(session_list)):
        
        # if starting a new session
        if session_list[s] is not curr_sesh:

            if s is not 0: # if not first line

                # write the old table
                if table_sorting == "tone":
                    tl_index = write_table(curr_table, sheet_total, tl_index, averaging = False, sorting = "tone") 
                    tl_index_avg = write_table(curr_table, sheet_average, tl_index_avg, averaging = True, sorting = "tone") 
                if table_sorting == "word":
                    tl_index = write_table(curr_table, sheet_total, tl_index, averaging = False, sorting = "word") 
                    tl_index_avg = write_table(curr_table, sheet_average, tl_index_avg, averaging = True, sorting = "word") 

                assert tl_index_avg == tl_index
            
            # if it is a new participant, make a new sheet
            if participant_list[s] is not curr_participant:
                print 'new sheet: ', participant_list[s]
                curr_participant = participant_list[s]
                sheet_total = wb.add_sheet(curr_participant + '_total')
                sheet_average = wb.add_sheet(curr_participant + '_average')
                tl_index = 0
                tl_index_avg = 0

            # make new table for new session
            curr_sesh = session_list[s]
            curr_table = Table(curr_sesh)

        
        # add word and value to accumulating table for this session
        if measuring == "target":
            value = col_dict["Total MWCM_Starget"][s+1]
        elif (measuring == "actual") | (measuring == "production"):
            value = col_dict["Total MWCM_Sactual"][s+1] 
        curr_table.add_word( Word(col_dict,s+1,tone=use_tone), value)
        
        
    # write the last table
    if table_sorting == "tone":
        write_table(curr_table, sheet_total, tl_index, averaging = False, sorting = "tone")
        write_table(curr_table, sheet_average, tl_index, averaging = True, sorting = "tone")
    if table_sorting == "word":
        write_table(curr_table, sheet_total, tl_index, averaging = False, sorting = "word") 
        write_table(curr_table, sheet_average, tl_index, averaging = True, sorting = "word") 
    
    # write the excel file
    if table_sorting == "tone": 
        if (use_tone == "target") & (measuring == "target"):
            wb.save('../output/MWCM_Starget_ToneCategory.xls')
        elif (use_tone == "target") & (measuring == "actual"):
            wb.save('../output/MWCM_Sactual_ToneCategory.xls')

    if table_sorting == "word": 
        if (use_tone == "target") & (measuring == "target"):
            wb.save('../output/MWCM_Starget_WordType.xls')
        elif (use_tone == "target") & (measuring == "actual"):
            wb.save('../output/MWCM_Sactual_WordType.xls')

def generate_table(use_session_list,sheet_name='gen_table', col_name = "Total MWCM_Starget", use_tone="target"):

    col_dict = ColumnDictionary('../output/Coding_output_words.xls')
    session_list = col_dict["Session"][1:]
    wb_tone = xlwt.Workbook()
    wb_word = xlwt.Workbook()

    #  make a new sheet
    sheet_total_tone = wb_tone.add_sheet(sheet_name + '_total')
    sheet_average_tone = wb_tone.add_sheet(sheet_name + '_average')

    sheet_total_word = wb_word.add_sheet(sheet_name + '_total')
    sheet_average_word = wb_word.add_sheet(sheet_name + '_average')

    # make new table for new session
    table = Table('table')

    for s in session_list:
        if s in use_session_list:

            # add word and value to accumulating table for this session
            value = col_dict[col_name][s+1]
            table.add_word( Word(col_dict,s+1,tone=use_tone), value)

    # write the last table
    write_table(table, sheet_total_tone, 0, averaging = False, sorting = "tone")
    write_table(table, sheet_average_tone, 0, averaging = True, sorting = "tone")
    write_table(table, sheet_total_word, 0, averaging = False, sorting = "word") 
    write_table(table, sheet_average_word, 0, averaging = True, sorting = "word") 

    # TODO cleanup excel write
# write the excel file
    if table_sorting == "tone": 
        if (use_tone == "target") & (measuring == "target"):
            wb.save('../output/MWCM_Starget_ToneCategory.xls')
        elif (use_tone == "target") & (measuring == "actual"):
            wb.save('../output/MWCM_Sactual_ToneCategory.xls')

    if table_sorting == "word": 
        if (use_tone == "target") & (measuring == "target"):
            wb.save('../output/MWCM_Starget_WordType.xls')
        elif (use_tone == "target") & (measuring == "actual"):
            wb.save('../output/MWCM_Sactual_WordType.xls')

if __name__ == "__main__":

    print 'MWCM value tables'
    MWCM(use_tone="target",measuring="actual", table_sorting="tone")
    MWCM(use_tone="target",measuring="target", table_sorting="tone")

    print 'MWCM word tables'
    MWCM(use_tone="target",measuring="actual",table_sorting="word")
    MWCM(use_tone="target",measuring="target",table_sorting="word")

    print 'Performing frequency analysis'
    FrequencyAnalysis(use_tone="target", table_sorting = "tone")
    FrequencyAnalysis(use_tone="production", table_sorting = "tone")
    FrequencyAnalysis(use_tone="target", table_sorting = "word")
    FrequencyAnalysis(use_tone="production", table_sorting = "word")
                

