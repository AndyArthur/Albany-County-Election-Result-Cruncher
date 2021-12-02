from functools import cmp_to_key
import re
import pandas as pd
import os
import datetime
import glob
import traceback

# number to excel column letter
def letter(colNum):
    import math
    
    if colNum > 26:
        return chr(ord('@')+math.floor(colNum/26))+chr(ord('@')+colNum%26)
    else:
        return chr(ord('@')+colNum)

enroll = pd.concat(pd.read_excel(f, header=4) for f in 
                ['/home/andy/enroll/feb20-Enrollment/'+x+'ED_feb20.xlsx' \
                 for x in ['Bronx','Kings','Queens','NewYork','Richmond']])

enroll = enroll[(enroll['STATUS']=='Active')] #active only
enroll=enroll.dropna() #drop totals

# Add Feb ADED code
enroll.index = enroll['ELECTION DIST'].str[-5:]
enroll = enroll.iloc[:,3:-2] # only keep party enroll

nenroll = pd.concat(pd.read_excel(f, header=4) for f in 
                ['/home/andy/enroll/nov20-Enrollment/'+x+'ED_nov20.xlsx' \
                 for x in ['Bronx','Kings','Queens','NewYork','Richmond']])

nenroll = nenroll[(nenroll['STATUS']=='Active')] #active only
nenroll=nenroll.dropna() #drop totals

# Add Nov ADED code
nenroll.index = nenroll['ELECTION DIST'].str[-5:]
nenroll = nenroll['TOTAL'] 

# bring february party enrollments together with november party enrollments
enroll=enroll.join(nenroll)

# start going through each race

for file in glob.glob('/home/andy/Desktop/2020-results/*csv'):
    try:

        race = os.path.basename(file)[11:-12]

        df = pd.read_csv(file, header=None)

        columns = list(df.iloc[0,0:11])
        df=df.iloc[:,11:]
        df.columns = columns

        pf = df.pivot(index=['AD','ED'],columns='Unit Name',values='Tally').fillna(0)

        # this line is dangerous and results should be checked for missing values!
        pf = pf.apply(pd.to_numeric, errors='coerce')

        # total ballots are everything but those with parties and scattered
        #totals = [x for x in list(pf.columns) if '(' not in x]
        #totals.remove('Scattered

        totals = ['Public Counter', 'Manually Counted Emergency',
                  'Absentee / Military','Federal','Special Presidential','Affidavit']

        totals = [x for x in totals if x in set(list(pf.columns))]


        pf['Ballot'] = pf[totals].sum(axis=1)
        pf.drop(totals, axis=1, inplace=True)
        pf.rename(columns={'Scattered': 'Other'}, inplace=True)

        # custom party sorting function to get
        # candidates in order like the board of election does
        def orderCol(x,y):
            # always first and list
            if x == 'Ballot' or y == 'Ballot':
                return -1
            if x == 'Other' or y == 'Other':
                return 1

            # get party name
            try: 
                partyX = re.findall('\((.*?)\)', x)[0]
                partyY = re.findall('\((.*?)\)', y)[0]

                # dictionary of parties by order
                pOrd = {'Democratic': 1, 'Republican': 2, 'Conservative': 3, 
                        'Green': 4, 'Working Families': 5, 'Independence': 6, 
                       'Womens Equality': 7, 'Libertarian': 8,'Reform': 9
                       }

                if pOrd[partyX] < pOrd[partyY]:
                    return -1

                if pOrd[partyX] > pOrd[partyY]:
                    return 1
            except:
                # when a candidate lacks a known party, such a party for the day, sent to end of list
                return 1

            return 0

        pf=pf[sorted(pf.columns, key=cmp_to_key(orderCol))]

        # make sure Other is after all others
        if 'Other' in pf.columns:
            pf.insert(len(pf.columns)-1, 'Other', pf.pop('Other'))

        pf.reset_index(inplace=True)

        ew = pd.ExcelWriter('/home/andy/Desktop/2020-output/2020 '+race+'.xlsx')

        for ad in pf['AD'].unique():

            raceStr = 'AD '+str(ad)
            af = pf[pf['AD']==ad]

            af.reset_index(inplace=True, drop=True)

            af.insert(2, 'Enrollment', 0 )
            af.insert(4, 'Blanks', af['Ballot']-af.iloc[:,4:].sum(axis=1))
            af.insert(5, 'Canvas', '=D'+(af.index+6).map(str)+'-E'+(af.index+6).map(str))  
            af.insert(6, 'TO %', '=(D'+(af.index+6).map(str)+'/C'+(af.index+6).map(str)+')')
            af.insert(7, 'DO %', '=(E'+(af.index+6).map(str)+'/D'+(af.index+6).map(str)+')')

            # create check columns
            af['CHECK'] = '=D'+(af.index+6).map(str)+'-E'+(af.index+6).map(str)
            af['CHECK %'] = '=0'

            # add percent columns
            for i, col in enumerate(af.columns[8:-2]): 
                try:
                    # check columns
                    if (len(af.columns)-i-2) > 0:
                        # add to check column
                        af['CHECK'] += '-'+letter(i+i+9)+(af.index+6).map(str)

                        # add check percent
                        af['CHECK %'] += '+'+letter(i+i+10)+(af.index+6).map(str)

                    # add percent
                    af.insert(i+i+9, col+' %', '=('+letter(i+i+9)+(af.index+6).map(str)+'/F'+(af.index+6).map(str)+')') 
                except Exception as e:
                    print(e)
                    pass

            # add in enrollment data, first by creating a key
            af['ADstr'] = af['ED'].map(str)
            af['ADstr'] = af['AD'].map(str)+af['ADstr'].str.zfill(3)    

            af.index = af['ADstr']
            af=af.join(enroll)
            
            # calculate type of enrollment needed -- for primaries, etc
            if (re.search('Democratic',race)):
                af['Enrollment'] = af['DEM']
 
            elif (re.search('Republican',race)):
                af['Enrollment'] = af['REP']
 
            elif (re.search('Conservative',race)):
                af['Enrollment'] = af['CON']

            elif (re.search('Working Families',race)):
                af['Enrollment'] = af['WFP']
 
            else:
                af['Enrollment'] = af['TOTAL']
            
            af.drop(columns=['ADstr'],inplace=True)
            af.drop(columns=enroll.columns,inplace=True)
            
            # disble header, manually write, as pre-defined headers can't be formatted
            af.to_excel(ew,sheet_name=raceStr, index=False, startrow=4)

            canForm = ew.book.add_format({'text_wrap': 0, 'font_name': 'Arial','font_size': 9, 'italic': True,
                                           'valign': 'vcenter', 'align': 'center'})
            headForm = ew.book.add_format({'text_wrap': 1, 'font_name': 'Arial','font_size': 9, 'bold': True , 
                                           'valign': 'vcenter', 'align': 'center'})

            rhForm = ew.book.add_format({'font_name': 'Arial','font_size': 9, 'bold': True , 
                                           'valign': 'vcenter', 'align': 'left'})

            for colnum, value in enumerate(af.columns.values):

                # find candidate names if they exist
                match = re.findall("(.*?) \((.*?)\)(.*?)$", value)

                if match:
                    # party name abbreviation and percent if exists
                    party = match[0][1][0:3].upper()
                    candidate = match[0][0]

                    if match[0][2]:
                        party += match[0][2]
                    else: # not a percent, so write candidate name
                        ew.sheets[raceStr].merge_range(letter(colnum+1)+'4:'+letter(colnum+2)+'4', candidate, canForm)

                    ew.sheets[raceStr].write(4, colnum, party, headForm)
                else:
                    ew.sheets[raceStr].write(4, colnum, value, headForm)

            # other header rows    
            ew.sheets[raceStr].merge_range('A1:H1', 'Assembly District '+str(ad), rhForm)
            ew.sheets[raceStr].merge_range('A2:H2', '2020 '+race, rhForm)
            
            if (re.search('Democratic',race)):
                ew.sheets[raceStr].merge_range('A3:H3', 'February DEM 2020 Active Enrollments', rhForm)
                ew.sheets[raceStr].set_row(4,40) # longer rows for primaries to fit names

            elif (re.search('Republican',race)):
                ew.sheets[raceStr].merge_range('A3:H3', 'February REP 2020 Active Enrollments', rhForm)
                ew.sheets[raceStr].set_row(4,40) # longer rows for primaries to fit names

            elif (re.search('Conservative',race)):
                ew.sheets[raceStr].merge_range('A3:H3', 'February CON 2020 Active Enrollments', rhForm)
                ew.sheets[raceStr].set_row(4,40) # longer rows for primaries to fit names

            elif (re.search('Working Families',race)):
                ew.sheets[raceStr].merge_range('A3:H3', 'February WFP 2020 Active Enrollments', rhForm)
                ew.sheets[raceStr].set_row(4,40) # longer rows for primaries to fit names
 
            else:
                ew.sheets[raceStr].merge_range('A3:H3', 'November 2020 Active Enrollments', rhForm)


            # set column width for all columns to 20, freeze panes
            bodyForm = ew.book.add_format({'text_wrap': 1, 'font_name': 'Arial','font_size': 9, 'num_format': '#,##0', 'valign': 'vcenter', 'align': 'right'})
            bodyPerForm = ew.book.add_format({'text_wrap': 1, 'font_name': 'Arial','font_size': 9, 'num_format': '0.0%', 'valign': 'vcenter', 'align': 'right'})
            muniForm = ew.book.add_format({'text_wrap': 1, 'font_name': 'Arial','font_size': 9, 'valign': 'vcenter', 'align': 'center'})

            ew.sheets[raceStr].set_column('A1:B9999', 4, muniForm)    
            ew.sheets[raceStr].set_column('C1:F9999', 10, bodyForm)
            ew.sheets[raceStr].set_column('G1:I9999', 10, bodyPerForm)

            for colNum in range(9,40):
                if colNum%2:
                    ew.sheets[raceStr].set_column(letter(colNum)+'1:'+letter(colNum)+'9999', 10, bodyForm)
                else:
                    ew.sheets[raceStr].set_column(letter(colNum)+'1:'+letter(colNum)+'9999', 10, bodyPerForm)

            ew.sheets[raceStr].freeze_panes('C6')
            ew.sheets[raceStr].set_landscape()
            ew.sheets[raceStr].set_paper(5)
            #ew.sheets[raceStr].set_header('&C Assembly District '+str(ad)+' - 2020 '+race)
            ew.sheets[raceStr].set_footer('&LCrunched by: NYC Number Cruncher.py on '  + \
                                          datetime.date.today().strftime('%-m/%-d/%Y, ') +
                                          'Checked by: ' +  \
                                          '&RPage &P of &N')
            ew.sheets[raceStr].repeat_rows(0,4)
            ew.sheets[raceStr].repeat_columns(0,1)
            ew.sheets[raceStr].fit_to_pages(1,0)


        ew.save()
        
        print('Race '+race+' was successfully processed and saved.')
        
    except:
        print('File '+race+' was unable to processed due to ')
        traceback.print_exc()
