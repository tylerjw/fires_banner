'''
Parses XLS files for Firse Banner program

Author: Tyler Weaver
'''
from xlrd import open_workbook, cellname
import re

def find_in_sheet(sheet, text):
    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            if sheet.cell(row_index,col_index).value == text:
                return (row_index, col_index)

def build_label_list(sheet, row):
    output = []
    for col in range(sheet.ncols):
        output.append(sheet.cell(row,col).value)
    return output

def short_cs(long_cs):
    end_re = re.compile('[A-Z][ ]?[0-9][0-9]')
    temp = end_re.search(long_cs)
    short_cs = ''
    if temp:
        end = temp.group()
        if(end[1] == ' '):
            end = end[0] + end[2:4]
        short_cs = long_cs[0] + end
    return short_cs

class ExcelFileOptions:
    def __init__(self):
        self.sh_name='SOTF-EAST PACE'
        self.label_row=1
        self.team_label='3/3 ODA'
        self.pfreq_label='P FREQ'
        self.afreq_label='A FREQ'
        self.jtar_label='JTAR #'
        self.ignore_label=['LATE DEPLOY']
        self.jtac_label='JTAC C/S' 
        self.jtar_prefix='KEG'

def build_fires_dicts(filename, file_options=ExcelFileOptions()):
    book = open_workbook(filename)
    sheet = book.sheet_by_name(file_options.sh_name)
    label_list = build_label_list(sheet, file_options.label_row)
    jtac_dict = {}
    team_dict = {}
    
    # build dictionary
    team_col = label_list.index(file_options.team_label)
    full_cs_col = label_list.index(file_options.jtac_label)
    pfreq_col = label_list.index(file_options.pfreq_label)+1 #get the NAME
    afreq_col = label_list.index(file_options.afreq_label) # test for TM INT
    jtar_col = label_list.index(file_options.jtar_label)
    team_re = re.compile('[0-9]{4,4}')
    jtac_sp_re = re.compile('[A-Z][A-Z] [0-9][0-9]')
    
    for row in range(file_options.label_row+1, sheet.nrows):
        jtac_cs_full = str(sheet.cell(row,full_cs_col).value).strip()
        jtac_cs = short_cs(jtac_cs_full)
        team_str = str(sheet.cell(row,team_col).value).strip()
        if jtac_cs != '' and jtac_cs_full not in file_options.ignore_label:
            #condition the input
            if(team_re.search(team_str)):
                team_str = team_str[0:4]
            if(jtac_sp_re.search(jtac_cs)):
                jtac_cs = jtac_cs[0:2] + jtac_cs[3:5]

            #team and jtac c/s
            if team_str != '':
                if team_dict.has_key(team_str):
                    try:
                        team_dict[team_str].index(jtac_cs_full)
                    except:
                        team_dict[team_str].append(jtac_cs_full)
                else:
                    team_dict[team_str] = [jtac_cs_full]

            if not jtac_dict.has_key(jtac_cs_full):
                jtac_dict[jtac_cs_full] = {}
                jtac_dict[jtac_cs_full]['JTAC CS'] = jtac_cs
            if jtac_dict[jtac_cs_full].has_key('team'):
                try:
                    jtac_dict[jtac_cs_full]['team'].index(team_str)
                except:
                    jtac_dict[jtac_cs_full]['team'].append(team_str)
            else:
                jtac_dict[jtac_cs_full]['team'] = [team_str]
            
            #print "%d: %s / %s"%(row,jtac_cs,team_str)

            #p freq name (9)
            pfreq_str = short_cs(str(sheet.cell(row,pfreq_col).value).strip())
            jtac_dict[jtac_cs_full]['P FREQ'] = pfreq_str

            #a freq + name (10, 11)
            afreq_str = str(sheet.cell(row,afreq_col).value).strip()
            afreq_name = str(sheet.cell(row,afreq_col+1).value).strip()
            if re.search('TM INT', afreq_str):
                afreq_str += ' ' + afreq_name
            elif re.search('TM INT', afreq_name):
                afreq_name += ' ' + afreq_str
                afreq_str = afreq_name
            else:
                afreq_str = short_cs(afreq_name)

            jtac_dict[jtac_cs_full]['A FREQ'] = afreq_str

            #JTAR# (KEG###)
            jtar_str = file_options.jtar_prefix + str(sheet.cell(row,jtar_col).value).strip()[0:3]
            jtac_dict[jtac_cs_full]['JTAR'] = jtar_str
            
    return (jtac_dict, team_dict)
    
if __name__ == '__main__':
    (jtacs, teams) = build_fires_dicts("documents\\SOTF EAST TADM 17MAY.xls")
    
    for key in teams:
        print key,
        print teams[key]

    print
    for key in jtacs:
        print key,
        print jtacs[key]
        
