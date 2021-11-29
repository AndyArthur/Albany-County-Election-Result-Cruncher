from io import StringIO
from styleframe import StyleFrame
import re
import pandas as pd
import numpy as np

# paths
# electionResultText = '/home/andy/Desktop/2021GeneralRecanvass.txt'
# enrollmentXLS = '/home/andy/enroll/nov21-Enrollment/AlbanyED_nov21.xlsx' 
# outputPath = '/tmp/2021_albany_county_races.xlsx'

# electionResultText = '/home/andy/Documents/GIS.Data/election.districts/2020GeneralRecanvass.txt'
# enrollmentXLS = '/home/andy/enroll/nov20-Enrollment/AlbanyED_nov20.xlsx' 
# outputPath = '/tmp/2020_albany_county_races.xlsx'

electionResultText = '/home/andy/Documents/GIS.Data/election.districts/2019'
enrollmentXLS = '/home/andy/enroll/nov19-Enrollment/AlbanyED_nov19.xlsx' 
outputPath = '/tmp/2018_albany_county_races.xlsx'

# number to excel column letter
def letter(colNum):
    import math
    
    if colNum > 26:
        return chr(ord('@')+math.floor(colNum/26))+chr(ord('@')+colNum%26)
    else:
        return chr(ord('@')+colNum)

pd.set_option("display.max_rows", None)
pd.set_option("display.max_columns", None) # show everything when previewing

with open(electionResultText) as f:
    data = f.read()

enroll = pd.read_excel(enrollmentXLS, header=4)
enroll = enroll[(enroll['STATUS']=='Active')][['ELECTION DIST', 'TOTAL']].dropna()
enroll['ELECTION DIST'] = enroll['ELECTION DIST'].astype(str).str.replace('  ',' ')
enroll['ELECTION DIST']=enroll['ELECTION DIST'].str.strip()
enroll['Enrollment'] = enroll['TOTAL']
enroll.drop(labels=['TOTAL'], axis=1, inplace=True)

# array containing election result dataframes
er = {}
    
# split each race which is divided by ten or more equal signs
for raceData in re.split('={10,}',data):

    # blank out old data frame
    df = None

    # split lines
    rows = raceData.split('\n')

    race = ""
    candidates = {}
    startAt = 0
    
    for i, line in enumerate(rows):
        # find race name
        if re.search('VOTES\s*?PERCENT', line):
            race = rows[i+1]
        
        race = race.replace('  ',' ')
        race = race.rstrip()
        #print(race)
        
        # find candidate names
        for result in re.findall('(\d\d)\s*?=\s*?(\w.*?)\d', line):
            candidates[int(result[0])] = result[1].rstrip()
       
        # find start at location for CSV reader
        if re.findall('-{5,}', line):
            startAt = i+3
            break
    
    # skips enrollment stats, as the don't have candidates
    if not candidates or not 2 in candidates:
        continue
    
    if not race:
        continue
        
    #print(race)
    
    df = pd.read_csv(
        StringIO(raceData),
        header=None,
        skiprows=startAt,
        sep='(?<=\d)\s{1,}(?=\d)',
        engine='python',
        on_bad_lines='warn')
    
    df=df[df[0].str.contains('^\d{4}').fillna(False)] # ONLY ROWS STARTING WITH 4-DIGIT ED CODE
    
    df.reset_index(drop=True, inplace=True) # reset index so we can use it in formulas

    df=df.rename(candidates,axis=1) # rename columns 
    
    df.iloc[:,1:]=df.iloc[:,1:].apply(pd.to_numeric) # make sure all columns are numeric 
    df.iloc[:,1:]=df.iloc[:,1:].convert_dtypes('int32') # cast columns to int32

    # crunching
    df.insert(1, 'Ballot', df.iloc[:,1:].convert_dtypes('int32').sum(axis=1)) # add total column
    
    df.insert(2, 'Blanks', df['OVER VOTES'].convert_dtypes('int32')) # add blank
    df.drop(labels=['OVER VOTES'], axis=1, inplace=True)

    df['Blanks']+=df['UNDER VOTES'].convert_dtypes('int32') # add under votes
    df.drop(labels=['UNDER VOTES'], axis=1, inplace=True)

    df.insert(3, 'Canvas', '=F'+(df.index+2).map(str)+'-G'+(df.index+2).map(str))  
    df.insert(4, 'TO %', '=(F'+(df.index+2).map(str)+'/E'+(df.index+2).map(str)+')')
    df.insert(5, 'DO %', '=(G'+(df.index+2).map(str)+'/F'+(df.index+2).map(str)+')')
        
    # create check columns
    df['CHECK'] = '=F'+(df.index+2).map(str)+'-G'+(df.index+2).map(str)
    df['CHECK %'] = '=0'
    
    # add percent columns
    for i, col in enumerate(df.columns[6:-2]): 
        try:
            # check columns
            if (len(df.columns)-i-2) > 0:
                # add to check column
                df['CHECK'] += '-'+letter(i+i+11)+(df.index+2).map(str)

                # add check percent
                df['CHECK %'] += '+'+letter(i+i+12)+(df.index+2).map(str)

            # add percent
            df.insert(i+i+7, col+' %', '=('+letter(i+i+11)+(df.index+2).map(str)+'/H'+(df.index+2).map(str)+')') 
        except:
            pass

    # add T W E Columns, drop combined field
    df[['ED Code','Municipality','String','Ward','ED']]=df[0].str.extract('(\d\d\d\d)\s*?(.*?)(\s*?WARD\s*(\d*))?\s*?ED\s*(\d*)')

    # temporary string used for data merge with enrollments
    df['ELECTION DIST'] = df['Municipality'].str.strip().replace('  ',' ') + \
    ' '+df['Ward'].fillna(0).astype(str).str.zfill(3)+df['ED'].str.zfill(3).astype('str')

    # merge on ED Key
    df=df.merge(enroll, on='ELECTION DIST', how='left')
    df=df.drop(0, axis=1)
    df=df.drop(labels=['ELECTION DIST','String'],axis=1)
    
    # move columns to proper order
    cols = list(df.columns)
    df = df[cols[-5:]+cols[:-5]]
    
    df['Municipality']=df['Municipality'].str.title()
    df['Municipality'] = df['Municipality'].str.strip().replace('  ',' ') # remove whitespace around muni column
    
    # array with election result dataframes
    er[race]=df.fillna(0)
    
# write file
ew = pd.ExcelWriter(outputPath)

for i, race in enumerate(er):
    raceStr = str(i+1)+' '+race
    if (len(race)>28):
        raceStr = str(i+1)+' '+race[:10] + '...' + race[-10:]
    
    # disble header, manually write, as pre-defined headers can't be formatted
    er[race].fillna(0).to_excel(ew,sheet_name=raceStr, index=False)
    
    headForm = ew.book.add_format({'text_wrap': 1, 'font_family': 'Arial', 'bold': True , 
                                   'valign': 'vcenter', 'align': 'center', 'bg_color': '#CCCCCC'})
    for colnum, value in enumerate(er[race].columns.values):
        ew.sheets[raceStr].write(0, colnum, value, headForm)
    
    ew.sheets[raceStr].set_row(0,50)
    
    # set column width for all columns to 20, freeze panes
    bodyForm = ew.book.add_format({'text_wrap': 1, 'font_family': 'Arial', 'num_format': '#,##0', 'valign': 'vcenter', 'align': 'right'})
    bodyPerForm = ew.book.add_format({'text_wrap': 1, 'font_family': 'Arial', 'num_format': '0.0%', 'valign': 'vcenter', 'align': 'right'})
    muniForm = ew.book.add_format({'text_wrap': 1, 'font_family': 'Arial', 'valign': 'vcenter', 'align': 'center'})
    ew.sheets[raceStr].set_column('A1:A9999', 6, bodyForm)
    ew.sheets[raceStr].set_column('B1:B9999', 15, muniForm)
    ew.sheets[raceStr].set_column('C1:D9999', 4, bodyForm)
    ew.sheets[raceStr].set_column('E1:H9999', 10, bodyForm)
    ew.sheets[raceStr].set_column('I1:J9999', 10, bodyPerForm)
    
    for colNum in range(11,40):
        if colNum%2:
            ew.sheets[raceStr].set_column(letter(colNum)+'1:'+letter(colNum)+'9999', 12, bodyForm)
        else:
            ew.sheets[raceStr].set_column(letter(colNum)+'1:'+letter(colNum)+'9999', 12, bodyPerForm)

    ew.sheets[raceStr].freeze_panes('E2')
        
ew.save()
