#!/usr/bin/env python

"""convertNASReports.py: read in heap autopay txt file file and export as single record text file (outAutopay.txt) """

__author__   = "Frederick Sulkowski"
__email__   = "frederick.sulkowski@erie.gov"


from argparse import ArgumentDefaultsHelpFormatter
from ftplib import ftpcp
from turtle import goto


def main():

    import os
    #import numpy    
    #import pandas as pd

    # read files in current directory
    for f_name in os.listdir('.'):
        if (f_name.startswith('001-X-WMMRELNA-14') and f_name.endswith('.txt')) or (f_name.startswith('001-X-WMMRRJNA-14') and f_name.endswith('.txt')) :
            wmmrelna(f_name)
        if f_name.startswith('001-X-WMMREXNA-14') and f_name.endswith('.txt') :
            wmmrexna(f_name)
        if f_name.startswith('001-X-WM4DELNA-14') and f_name.endswith('.txt') :
            wm4delna(f_name)
        if f_name.startswith('001-X-WM4DEXNA-14') and f_name.endswith('.txt') :
            wm4dexna(f_name)
        if f_name.startswith('001-X-14-WMSCSDX5') and f_name.endswith('.txt') :
            wmscsdx5(f_name)
        if (f_name.startswith('001-X-14-WMSBF630') and f_name.endswith('.txt')) or (f_name.startswith('001-X-14-WMSBP630') and f_name.endswith('.txt')) :
            wmsbf630(f_name)
        if (f_name.startswith('001-X-14-WMSBF650') and f_name.endswith('.txt')) or (f_name.startswith('001-X-14-WMSBP650') and f_name.endswith('.txt')) :
            wmsbf650(f_name)
        if (f_name.startswith('001-X-14-WMSBF650') and f_name.endswith('.txt')) or (f_name.startswith('001-X-14-WMSBP650') and f_name.endswith('.txt')) :
            wmsbf650(f_name)
        if (f_name.startswith('001-X-14-WMSA5257') and f_name.endswith('.txt')) :
            wmsa5257(f_name)
        if (f_name.startswith('001-X-14-WMSC1025') and f_name.endswith('.txt')) or (f_name.startswith('001-X-14-WMSC1040') and f_name.endswith('.txt')) :
            wmsc1025(f_name)
        if (f_name.startswith('001-X-14-WMSC1026') and f_name.endswith('.txt')) or (f_name.startswith('001-X-14-WMSC1041') and f_name.endswith('.txt')) :
            wmsc1026(f_name)
        if (f_name.startswith('001-X-14-WMSC2027') and f_name.endswith('.txt')) :
            wmsc2027(f_name)            
        if (f_name.startswith('001-X-14-WMSBHIGH') and f_name.endswith('.txt')) :
            wmsbhigh(f_name)            
        if (f_name.startswith('001-X-14-WMSC4210') and f_name.endswith('.txt')) :
            wmsc4210(f_name)            

#===================================================================================================

def wmmrelna(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'

    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    startDetail = True

    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()

    out_str = ("OFFICE\t" + "UNIT\t" + "WORKER\t" + "CASE NAME\t" + "CASE NUMBER\t" + "OLD DEFCT\t" + "NEW DEFCT\t" + "NET CHNG\t" +
                "SPEC ALERT\t" + "OLD ETLMT\t" + "NEW ETLMT\t" + "NET CHNG\t" + "REJ/EXCEPT")

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if line.startswith('FOR MASS REBUDGETING'):
            startDetail = False

        if line.startswith('DISTRICT') and startDetail == True:
            if 'WMMRRJNA' in f_name:
                office = record[54:59].strip() + '\t'
                unit = record[65:70].strip() + '\t'
                worker = record[80:86].strip() + '\t'
            else:
                office = record[44:49].strip() + '\t'
                unit = record[56:61].strip() + '\t'
                worker = record[72:78].strip() + '\t'

        if (not line.startswith('WMRIQ5') and not line.startswith('DISTRICT') and not line.startswith('OLD DEFCT') 
            and not line.startswith('CASE NAME') and not line.startswith(escape) 
            and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseName = record[26:47].strip() + '\t'
                caseNumber = record[48:58].strip() + '\t'
                oldDefct = record[60:67].strip() + '\t'
                newDefct = record[72:78].strip() + '\t'
                netChng1 = record[82:89].strip() + '\t'
                specAlert = record[94:].strip() + '\t'

            if countLines == 2:
                oldEtlmt = record[60:67].strip() + '\t'
                newEtlmt = record[71:78].strip() + '\t'
                netChng2 = record[82:89].strip() + '\t'
                rejExcept = record[94:]
                endPos = rejExcept.find(carriage)
                rejExcept = rejExcept[0:endPos].strip()
                out_str = (office + unit + worker + caseName + caseNumber + oldDefct + newDefct + netChng1 + specAlert + 
                           oldEtlmt + newEtlmt + netChng2 + rejExcept)

            # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0

    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)               
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")
    os.remove(outfile)            

    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    if 'WMMRRJNA' in f_name:
        messagebox.showinfo("Completed - WMMRRJNA", "Completed!  " + str(countWrites) + " records converted from " + f_name + "!")
    else:
        messagebox.showinfo("Completed - WMMRELNA", "Completed!  " + str(countWrites) + " records converted from " + f_name + "!")

    return

#===================================================================================================

def wmmrexna(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'

    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    startDetail = True

    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()

    out_str = ("OFFICE\t" + "UNIT\t" + "WORKER\t" + "CASE NAME\t" + "CASE NUMBER\t" + "EXCEPTED REASON")

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if line.startswith('FOR MASS REBUDGETING'):
            startDetail = False

        if line.startswith('DISTRICT') and startDetail == True:
            office = record[54:59].strip() + '\t'
            unit = record[65:70].strip() + '\t'
            worker = record[80:86].strip() + '\t'

        if (not line.startswith('WMRIQ4') and not line.startswith('DISTRICT') and not line.startswith('CASE NAME') 
            and not line.startswith(escape) and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseName = record[28:55].strip() + '\t'
                caseNumber = record[56:66].strip() + '\t'
                exceptedReason = record[84:].strip()
                out_str = (office + unit + worker + caseName + caseNumber + exceptedReason)

            # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0

    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)               
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")
    os.remove(outfile)            

    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMMREXNA", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wm4delna(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'

    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    startDetail = True

    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()

    out_str = ("OFFICE\t" + "UNIT\t" + "WORKER\t" + "CASE NAME\t" + "CASE NUMBER\t" + "MESSAGE #1\t" + "PREV 2\t" +
                "PREV 1\t" + "CUR IVD\t" + "NEW FS\t" + "CUR FA\t" + "SPEC ALERTS\t" + "MESSAGE #2\t" + "CUM 2\t" +
                "CUM 1\t" + "CUR CUM\t" + "OLD FS\t" +"PRE FA\t" + "ACTION CODE\t" + "MONTH OBL\t" + "EXMT 2\t" +
                "EXMT 1\t" + "EXMT CUR")

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if line.startswith('FOR MASS REBUDGETING'):
            startDetail = False

        if line.startswith('DISTRICT') and startDetail == True:
            office = record[47:52].strip() + '\t'
            unit = record[58:64].strip() + '\t'
            worker = record[75:81].strip() + '\t'

        if (not line.startswith('WMRIV2') and not line.startswith('DISTRICT') and not line.startswith('CASE NAME')
            and not line.startswith('CASE NUMBER') and not line.startswith('MONTH OBL') and not line.startswith(escape) 
            and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseName = record[9:39].strip() + '\t'
                message1 = record[39:50].strip() + '\t'
                prev2 = record[54:61].strip() + '\t'
                prev1 = record[65:72].strip() + '\t'
                curIvd = record[76:83].strip() + '\t'
                newFs = record[87:94].strip() + '\t'
                curFa = record[98:105].strip() + '\t'
                specAlerts = record[109:].strip() + '\t'
                
            if countLines == 2  :
                caseNumber = record[14:24].strip() + '\t'
                message2 = record[39:50].strip() + '\t'
                cum2 = record[54:61].strip() + '\t'
                cum1 = record[65:72].strip() + '\t'
                curCum = record[76:83].strip() + '\t'
                oldFs = record[87:94].strip() + '\t'
                preFa = record[98:105].strip() + '\t'
                actionCode = record[109:].strip() + '\t'

            if countLines == 3  :
                monthObl = record[39:50].strip() + '\t'
                exmt2 = record[54:61].strip() + '\t'
                exmt1 = record[65:72].strip() + '\t'
                exmtCur = record[76:83].strip()
                out_str = (office + unit + worker + caseName + caseNumber + message1 + prev2 + prev1 + curIvd + newFs +
                           curFa + specAlerts + message2 + cum2 + cum1 + curCum + oldFs + preFa + actionCode + monthObl +
                           exmt2 + exmt1 + exmtCur)

            # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0

    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)               
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")
    os.remove(outfile)            

    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WM4DELNA", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wm4dexna(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd
 
    escape = '\u001B'
    carriage = '\r'
    nullChar = '\u0000' 

    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    startDetail = True

    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()

    out_str = ("OFFICE\t" + "UNIT\t" + "WORKER\t" + "CASE NAME\t" + "CASE NUMBER\t" + "MESSAGE #1\t" + "PREV 2\t" +
                "PREV 1\t" + "CUR IVD\t" + "CUR FA\t" + "EXCEPTION REASON\t" + "MESSAGE #2\t" + "CUM 2\t" + "CUM 1\t" +
                "CUR CUM\t" + "PRE FA\t" +"ACTION REQUIRED\t" + "MONTH OBL\t" + "EXMT 2\t" + "EXMT 1\t" + "EXMT CUR")

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if line.startswith('FOR MASS REBUDGETING'):
            startDetail = False

        if line.startswith('DISTRICT') and startDetail == True:
            office = record[47:52].strip() + '\t'
            unit = record[58:64].strip() + '\t'
            worker = record[75:81].strip() + '\t'

        if (not line.startswith('WMRIV1') and not line.startswith('DISTRICT') and not line.startswith('CASE NAME')
            and not line.startswith('CASE NUMBER') and not line.startswith('MONTH OBL') and not line.startswith(escape) 
            and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseName = record[9:39].strip() + '\t'
                message1 = record[39:50].strip() + '\t'
                prev2 = record[54:61].strip() + '\t'
                prev1 = record[65:72].strip() + '\t'
                curIvd = record[76:83].strip() + '\t'
                curFa = record[87:93].strip() + '\t'
                exceptionReason = record[98:].replace(nullChar,' ').strip() + '\t'

            if countLines == 2  :
                caseNumber = record[14:24].strip() + '\t'
                message2 = record[39:50].strip() + '\t'
                cum2 = record[54:61].strip() + '\t'
                cum1 = record[65:72].strip() + '\t'
                curCum = record[76:83].strip() + '\t'
                preFa = record[87:93].strip() + '\t'
                actionRequired = record[98:].strip() + '\t'

            if countLines == 3  :
                monthObl = record[39:50].strip() + '\t'
                exmt2 = record[54:61].strip() + '\t'
                exmt1 = record[65:72].strip() + '\t'
                exmtCur = record[76:83].strip()
                out_str = (office + unit + worker + caseName + caseNumber + message1 + prev2 + prev1 + curIvd +
                           curFa + exceptionReason + message2 + cum2 + cum1 + curCum + preFa + actionRequired +
                           monthObl + exmt2 + exmt1 + exmtCur)

            # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0
               
    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)               
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")
    os.remove(outfile)            

    messagebox.showinfo("Completed - WM4DEXNA", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmscsdx5(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd
    
    escape = '\u001B'
    carriage = '\r'
    nullChar = '\u0000' 

    # setup count variables
    countLines = 0
    countWrites = 0
    statusNext = ' '

    outfile = "ec_" + f_name
    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfile):
        os.remove(outfile)
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()

    out_str = ("STATUS\t" + "CASE TYPE\t" + "SSN\t" + "LAST NAME\t" + "FIRST NAME\t" + "M I\t" + "CASE NUMBER\t" +
                "CIN\t" + "OLD ETMNT\t" + "NEW ETMNT\t" + "NEW ALMNT\t" + "F T\t" + "AUTH FROM\t" + "AUTH TO\t" + "TRANSACTION DISPOSITION\t" +
                "TRANSACTION DISPOSITION\t" + "TRANSACTION DISPOSITION\t" + "TRANSACTION DISPOSITION")

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if statusNext == 'Y' and not line.startswith(escape):
            status = record[0:64].strip() + '\t'
            statusNext = ' '

        if line.startswith('TYPE'):
            statusNext = 'Y'
        
        if (not line.startswith('DISTRICT') and not line.startswith('AUTO SDX') and not line.startswith('AUTOMATED')
            and not line.startswith('TRANSACTION CONTROL REPORT') and not line.startswith('CASE') 
            and not line.startswith('TYPE') and not line.startswith('EXCEPTION') and not line.startswith('ELIGIBLE')
            and not line.startswith('RFC') and not line.startswith('NON WORKER') 
            and not line.startswith(escape) and not line.startswith(carriage) and len(line) != 0) :

            if countLines == 3 :
                countLines = 0

            if countLines == 2 :
                recordCheck = record[0:20].strip()

                if recordCheck == "" :
                    transactionDisposition3 = record[26:60].strip() + '\t'
                    transactionDisposition4 = record[61:95].strip()
                else :
                    transactionDisposition3 = ' ' + '\t'
                    transactionDisposition4 = ' '
                    countLines = 0

                out_str = (status + caseType + ssn + lastName + firstName + mi + caseNumber + cin + oldEtmnt +
                    newEtmnt + newAlmnt + ft + authFrom + authTo + transactionDisposition1 + transactionDisposition2 +
                    transactionDisposition3 + transactionDisposition4) 
                           
            # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 

            countLines = countLines + 1

            if countLines == 1 :
                caseType = record[0:4].strip() + '\t'
                if record[7:8].strip() == 'd' :
                    record2 = '       d           ' + record[11:]
                    record = record2
                ssn = record[7:16].strip() + '\t'
                lastName = record[19:39].strip() + '\t'
                firstName = record[39:50].strip() + '\t'
                mi = record[50:51].strip() + '\t'
                caseNumber = record[54:64].strip() + '\t'
                cin = record[67:75].replace(nullChar,' ').strip() + '\t'
                oldEtmnt = record[78:86].strip() + '\t' 
                newEtmnt = record[89:97].strip() + '\t'
                newAlmnt = record[100:107].strip() + '\t'
                ft = record[110:111].strip() + '\t'
                authFrom = record[114:122].strip() + '\t'
                authTo  = record[122:130].strip() + '\t'

            if countLines == 2 :
                transactionDisposition1 = record[26:60].strip() + '\t'
                transactionDisposition2 = record[61:95].strip() + '\t'

    out_str = (status + caseType + ssn + lastName + firstName + mi + caseNumber + cin + oldEtmnt +
        newEtmnt + newAlmnt + ft + authFrom + authTo + transactionDisposition1 + transactionDisposition2 +
        transactionDisposition3 + transactionDisposition4)
                           
    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
        countWrites = countWrites +1 

    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WM4DEXNA", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsbf630(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd
    
    escape = '\u001B'
    carriage = '\r'
    nullChar = '\u0000' 

    # setup count variables
    countLines = 0
    countWrites = 0
    heapPerson = ' '

    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()

    out_str = ("CASE NUMBER\t" + "CASE NAME\t" + "LASTNAME\t" + "FIRSTNAME\t" + "MI\t" + "EXCEPTION REASON")

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if (not line.startswith('DISTRICT') and not line.startswith('LIST OF')
            and not line.startswith('CASE NUMBER') and not line.startswith(escape)) :
            # and not line.startswith(carriage) and len(line) != 0) :

                    countLines = countLines + 1
                    lengthOfLine = len(line)

            # format 3 lines of data into one            
                    if line[0] != '\"':
                        line = ("\"" + line + "\"")

                    heapPerson = heapPerson + line

                    if countLines == 1 and lengthOfLine < 118:
                        heapPerson = heapPerson[:116] + "   \""

                    if countLines == 1:
                        new_str = heapPerson.replace("\"", " ")

            # case number (caseNumber)
                        caseNumber = new_str[2:13].strip() + '\t'
            # case name (caseName)                
                        if "." in new_str[15:51]:
                            new_str = new_str[:15] + new_str[15:51].replace("."," ") + new_str[51:]

                        if ", " in new_str[15:51] or " ," in new_str[15:51]:
                            new_str = new_str.replace(", ",",")
                            new_str = new_str.replace(" ,",",")
                    
                        tempName = new_str[15:50].strip()
                        caseName = new_str[15:50].strip() + '\t'


#                        def get_name_parts(name):
                        comma_split = tempName.split(',')
                        tempLast = comma_split[0]
                        caseLast = str(tempLast) + '\t'

                        try:
                            tempFirst = comma_split[1].split(',')[0]
                            caseFirst = str(tempFirst) + '\t'
                        except IndexError:
                            caseFirst = ' \t'

                        try:
                            tempMI = comma_split[2].split(',')[0]
                            caseMI = str(tempMI) + '\t'
                        except IndexError:
                            caseMI = ' \t'

            # exception reason (exceptionReason)
                        exceptionReason = new_str[52:76].strip()


               
                        out_str = caseNumber + caseName + caseLast + caseFirst + caseMI + exceptionReason

            # write to results file                                
                        if out_str != ' ':
                            f = open(outfile,"a")
                            f.write(out_str + "\n")
                            f.close()

                            countWrites = countWrites +1 
                            countLines = 0
                            heapPerson = ' '

    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None,low_memory=False)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMSBF630", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsbf650(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd
    
    escape = '\u001B'
    carriage = '\r'
    nullChar = '\u0000' 

    # setup count variables
    countLines = 0
    countWrites = 0

    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    # read file
    file = open(f_name)
    lines = file.readlines()
    file.close()

    out_str = ("CASE NUMBER\t" + "CASE NAME\t" + "ADDRESS\t" + "CITY\t" + "STATE\t" + "ZIP\t" + "Local Office\t" + "UNIT\t" + 
                            "WORKER\t" + "ACT (Local Action Code)\t" + "PAY TYPE\t" + "Method\t" + "AMOUNT\t" + "IS (Issuance Code)\t" + 
                            "SC (Schedule Code)\t" + "PU (Pick Up Code)\t" + "PERIOD FROM\t" + "PERIOD TO\t" + "FUEL TYPE\t" +
                            "SPC CLM (Special Claim\t)" + "LOCAL USAGE\t" + "VP (Voucher Payment\t" + "HH SIZE\t" + "INCOME\t" +
                            "HEAP VENDOR ID\t" + "CUSTOMER ACCT NO\t" + "NX (Heap Nominal Benefit\t" + "CASE TYPE")

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()

        # setup variables
        heapPerson = ' '
    
        for line in lines:
            line = line.strip().upper()
            line = line.strip()

            if (not line.startswith('CASE NAME') and not line.startswith('DISTRICT') and not line.startswith('ADDRESS') 
                and not line.startswith('CASE NUMBER') and not line.startswith('RC132') and not line.startswith('* * *  DIS') 
                and not line.startswith('"*  TOTAL') and not line.startswith('*  TOTAL') and not line.startswith("\u001B")):

                countLines = countLines + 1
                lengthOfLine = len(line)

                # format 3 lines of data into one            
                if line[0] != '\"':
                    line = ("\"" + line + "\"")

                heapPerson = heapPerson + line

                if countLines == 1 and lengthOfLine < 118:
                    heapPerson = heapPerson[:116] + "   \""

    #            if countLines == 2 and heapPerson[158] == ' ':
    #                heapPerson = heapPerson[:158] + heapPerson[160:]

                if countLines == 3 and heapPerson[188] == '"':
                    heapPerson = heapPerson[:189] + '           ' + heapPerson[189:]

                if countLines == 3:
                    new_str = heapPerson.replace("\"", " ")
                    
                    caseNumber = new_str[2:14].strip() + '\t'
                    caseName = new_str[15:51].strip() + '\t'
                    caseAddress = new_str[121:157].strip() + '\t'
                    caseCity = new_str[201:217].strip() + '\t'
                    caseState = new_str[219:221].strip() + '\t'
                    caseZip = new_str[225:230].strip() + '\t'
                    localOffice = new_str[52:55].strip() + '\t'
                    caseUnit = new_str[158:163].strip() + '\t'
                    caseWorker = new_str[238:243].strip() + '\t'
                    localActionCode = new_str[60:61].strip() + '\t'
                    payType = new_str[64:66].strip() + '\t'
                    caseMethod = new_str[70:72].strip() + '\t'
                    caseAmount = new_str[73:82].strip() + '\t'
                    issuanceCode = new_str[84:85].strip() + '\t'
                    scheduleCode = new_str[87:87].strip() + '\t'
                    pickUpCode = new_str[90:91].strip() + '\t'
                    periodFrom = new_str[94:100].strip() + '\t'
                    periodTo = new_str[101:107].strip() + '\t'
                    fuelType = new_str[109:110].strip() + '\t'
                    spcClm = new_str[114:115].strip() + '\t'
                    localUsage = new_str[118:118].strip() + '\t'
                    voucherPayment = new_str[179:180].strip() + '\t'
                    hhSize = new_str[186:188].strip() + '\t'
                    caseIncome = new_str[189:199].strip() + '\t'
                    heapVendorId = new_str[250:257].strip() + '\t'
                    customerAcctNo = new_str[270:285].strip() + '\t'
                    heapNominalBenefit = ' ' + '\t'
                    caseType = new_str[309:311].strip()

                    out_str = (caseNumber + caseName + caseAddress + caseCity + caseState + caseZip + localOffice + caseUnit + caseWorker + 
                        localActionCode + payType + caseMethod + caseAmount + issuanceCode + scheduleCode + pickUpCode + periodFrom + periodTo + fuelType +
                        spcClm + localUsage + voucherPayment + hhSize + caseIncome + heapVendorId + customerAcctNo + heapNominalBenefit + caseType)

            # write to results file                                
                    if out_str != ' ':
                        f = open(outfile,"a")
                        f.write(out_str + "\n")
                        f.close()

                        countWrites = countWrites +1 
                        countLines = 0
                        heapPerson = ' '



    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None,low_memory=False)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMSBF630", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsa5257(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

#    import fnmatch
#    import time
#    import operator
    
#    from operator import itemgetter

###    timestr = time.strftime("%Y%m%d.%H%M%S")
    escape = '\u001B'
    carriage = '\r'
            
    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    
    # read file
    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()
            
    getVendorName = False
    startDetail = False

    out_str = ("LOCAL OFFICE\t" + "UNIT\t" + "WORKER\t" + "CASE NUMBER\t" + "CASE NAME\t" + "INDIVIDUAL NAME\t" + "CIN\t" +
        "RETURN DATE")

    print("f_name is " + f_name)

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()

    # setup variables
    heapPerson = ' '
    
    for line in lines:
        line = line.strip().upper()
        line = line.strip()
               
        if line.startswith('LOCAL OFFICE'):
            localOffice = line[15:18].strip() + '\t'
            unit = line[37:42].strip() + '\t'
            worker = line[61:66].strip() + '\t'
            startDetail = False

        if line.startswith('CASE NUMBER'):
                    startDetail = True

        if (not line.startswith('REPORT') and not line.startswith('*') and not line.startswith('QI-1') 
            and not line.startswith('PERIOD COVERED BY') and not line.startswith('REFERENCE') and not line.startswith('DISTRICT')
            and not line.startswith('LOCAL OFFICE') and not line.startswith('CASE NUMBER') and not line.startswith(escape) 
            and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseNumber = line[:10].strip() + '\t'
                caseName = line[21:57].strip() + '\t'
                individualName = line[59:97].strip() + '\t'
                cin = line[99:107].strip() + '\t'
                returnDate = line[117:125].strip()

                out_str = (localOffice + unit + worker + caseNumber + caseName + individualName + cin + returnDate)

    # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0
                    heapPerson = ' '
              
    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMSA5257", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsc1025(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'
            
    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    
    # read file
    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()
            
    getVendorName = False
    startDetail = False

    out_str = ("UNIT\t" + "LOCAL OFFICE\t" + "WORKER\t" + "CASE NUMBER\t" + "CASE NAME\t" + "Payee\t" + "CASE TYPE\t" +
        "PAY AMOUNT\t" + "PAY PERIOD")

    print("f_name is " + f_name)

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()

    # setup variables
    heapPerson = ' '
    
    for line in lines:
        line = line.strip().upper()
        line = line.strip()

        if line.startswith('DISTRICT'):
            unit = line[51:56].strip() + '\t'
            startDetail = False

        if line.startswith('LOCAL OFFICE'):
            localOffice = line[15:18].strip() + '\t'
            worker = line[53:58].strip() + '\t'
            startDetail = False

        if line.startswith('CASE NUMBER'):
                    startDetail = True

        if (not line.startswith('REPORT') and not line.startswith('*')
            and not line.startswith('MASS AUTHORIZATION') and not line.startswith('DISTRICT') 
            and not line.startswith('LOCAL OFFICE') and not line.startswith('CASE NUMBER')
            and not line.startswith('WMS REPORT') and not line.startswith('REFERENCE') 
            and not line.startswith('END OF REPORT') and not line.startswith('TOTAL') 
            and not line.startswith(escape) and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseNumber = line[:10].strip() + '\t'
                caseName = line[23:55].strip() + '\t'
                if '/PAYEE' in caseName :
                    caseName = caseName.replace('/PAYEE','')
                    payee = 'TRUE\t'
                else :
                    payee = '\t'
                caseType = line[56:58].strip() + '\t'
                payAmount = line[67:75].strip() + '\t'
                payPeriod = line[82:101].strip()

                out_str = (localOffice + unit + worker + caseNumber + caseName + payee + caseType + payAmount + payPeriod)

    # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0
                    heapPerson = ' '
              
    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMSC1025/1040", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsc1026(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'
            
    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    
    # read file
    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()
            
    getVendorName = False
    startDetail = False

    out_str = ("UNIT\t" + "LOCAL OFFICE\t" + "WORKER\t" + "CASE NUMBER\t" + "CASE NAME\t" + "Payee\t" + "CASE TYPE\t" +
        "EXCEPTION REASON")

    print("f_name is " + f_name)

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()

    # setup variables
    heapPerson = ' '
    
    for line in lines:
        line = line.strip().upper()
        line = line.strip()

        if line.startswith('DISTRICT'):
            unit = line[51:56].strip() + '\t'
            startDetail = False

        if line.startswith('LOCAL OFFICE'):
            localOffice = line[15:18].strip() + '\t'
            worker = line[53:58].strip() + '\t'
            startDetail = False

        if line.startswith('CASE NUMBER'):
                    startDetail = True

        if (not line.startswith('REPORT') and not line.startswith('*')
            and not line.startswith('MASS AUTHORIZATION') and not line.startswith('DISTRICT') 
            and not line.startswith('LOCAL OFFICE') and not line.startswith('CASE NUMBER')
            and not line.startswith('WMS REPORT') and not line.startswith('REFERENCE') 
            and not line.startswith('END OF REPORT') and not line.startswith('TOTAL') 
            and not line.startswith(escape) and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseNumber = line[:10].strip() + '\t'
                caseName = line[13:45].strip() + '\t'
                if '/PAYEE' in caseName :
                    caseName = caseName.replace('/PAYEE','')
                    payee = 'TRUE\t'
                else :
                    payee = '\t'
                caseType = line[46:48].strip() + '\t'
                exceptionReason = line[68:].strip()

                out_str = (localOffice + unit + worker + caseNumber + caseName + payee + caseType + exceptionReason)

    # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0
                    heapPerson = ' '
              
    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMSC1026/1041", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsc2027(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'
            
    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    
    # read file
    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

#    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    outfilexlsx = "WINR2027 MASS ISSUANCE OF P-EBT - ELIGIBLE REPORT [" + os.path.splitext(f_name[28:38])[0] + "].xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()
            
    getVendorName = False
    startDetail = False

    out_str = ("UNIT\t" + "LOCAL OFFICE\t" + "WORKER\t" + "CASE NUMBER\t" + "CASE NAME\t" + "Payee\t" + "CASE TYPE\t" +
        "PAY AMOUNT\t" + "PAY PERIOD FROM\t" + "PAY PERIOD TO\t" + "AVAIL DT")

    print("f_name is " + f_name)

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()

    # setup variables
    heapPerson = ' '
    
    for line in lines:
        line = line.strip().upper()
        line = line.strip()

        if line.startswith('DISTRICT'):
            unit = line[51:56].strip() + '\t'
            startDetail = False

        if line.startswith('LOCAL OFFICE'):
            localOffice = line[15:18].strip() + '\t'
            worker = line[53:58].strip() + '\t'
            startDetail = False

        if line.startswith('CASE NUMBER'):
                    startDetail = True

        if (not line.startswith('REPORT') and not line.startswith('*')
            and not line.startswith('MASS ISSUANCE') and not line.startswith('DISTRICT') 
            and not line.startswith('LOCAL OFFICE') and not line.startswith('CASE NUMBER')
            and not line.startswith('WMS REPORT') and not line.startswith('REFERENCE') 
            and not line.startswith('END OF REPORT') and not line.startswith('TOTAL') 
            and not line.startswith(escape) and not line.startswith(carriage) and len(line) != 0 and startDetail == True):

            countLines = countLines + 1

            if countLines == 1  :
                caseNumber = line[:10].strip() + '\t'
                caseName = line[23:55].strip() + '\t'
                if '/PAYEE' in caseName :
                    caseName = caseName.replace('/PAYEE','')
                    payee = 'TRUE\t'
                else :
                    payee = '\t'
                caseType = line[56:58].strip() + '\t'
                payAmount = line[67:75].strip() + '\t'
                payPeriodFrom = line[82:90].strip() + '\t'
                payPeriodTo = line[93:101].strip() + '\t'
                availDt = line[107:117].strip()

                out_str = (localOffice + unit + worker + caseNumber + caseName + payee + caseType + payAmount + payPeriodFrom
                    + payPeriodTo + availDt)

    # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0
                    heapPerson = ' '
              
    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMSC2027", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsbhigh(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'
            
    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    
    # read file
    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

#    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    outfilexlsx = "WMSBHIGH - RFI HIGH RISK LISTING [" + os.path.splitext(f_name[28:38])[0] + "].xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()
            
    getVendorName = False
    startDetail = False

    out_str = ("LOCAL OFFICE\t" + "UNIT\t" + "WORKER\t" + "CASE TYPE\t" + "CASE NUMBER\t" + "CASE NAME\t" + "Payee\t" + "INCOME\t" +
        "RESOURCE")

    print("f_name is " + f_name)

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if (not line.startswith('REPORT') and not line.startswith('*')
            and not line.startswith('DISTRICT') and not line.startswith('OFFICE') 
            and not line.startswith('RFI HIGH') and not line.startswith('UNRESOLVED')
            and not line.startswith('SORT SEQUENCE') and not line.startswith('PLEASE') 
            and not line.startswith('THE ATTACHED') and not line.startswith('THE UNRESOLVED')
            and not line.startswith('FOR QUESTIONS') and not line.startswith('ASSISTANCE')
            and not line.startswith('HITS  THAT') and not line.startswith('REVENUE')
            and not line.startswith('BUSINESS') and not line.startswith('45 DAYS')
            and not line.startswith('"HITS"') and not line.startswith('OF THE WRM.')
            and not line.startswith(escape) and not line.startswith(carriage) and len(line) != 0):

            countLines = countLines + 1

            if countLines == 1  :
                localOffice = record[5:8].strip() + '\t'
                unit = record[14:19].strip() + '\t'
                worker = record[24:29].strip() + '\t'
                caseType = record[36:38].strip() + '\t'
                caseNumber = record[52:62].strip() + '\t'
                caseName = record[66:96].strip() + '\t'
                if '/PAYEE' in caseName :
                    caseName = caseName.replace('/PAYEE','')
                    payee = 'TRUE\t'
                else :
                    payee = '\t'
                income = record[98:105].strip() + '\t'
                resource = record[110:120].strip()

                out_str = (localOffice + unit + worker + caseType + caseNumber + caseName + payee + income + resource)

    # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0
                    heapPerson = ' '
              
    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WMSC2027", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

def wmsc4210(f_name):

    from tkinter import messagebox
    import os
    import pandas as pd

    escape = '\u001B'
    carriage = '\r'
            
    # setup count variables
    countLines = 0
    countWrites = 0
    countFiles = 0
    
    # read file
    outfile = "ec_" + f_name
    if os.path.exists(outfile):
        os.remove(outfile)

#    outfilexlsx = "ec_" + os.path.splitext(f_name)[0] + ".xlsx"
    outfilexlsx = "WINR4210 HEAP AUTO-CLOSE EXCEPTION REPORT [" + os.path.splitext(f_name[28:38])[0] + "].xlsx"
    if os.path.exists(outfilexlsx):
        os.remove(outfilexlsx)

    file = open(f_name)
    lines = file.readlines()
    file.close()
            
    getVendorName = False
    startDetail = False

    out_str = ("LOCAL OFFICE\t" + "UNIT\t" + "WORKER\t" + "CASE NAME\t" + "CASE NUMBER\t" + "EXCEPTION REASON")

    print("f_name is " + f_name)

    # write to results file                                
    if out_str != ' ':
        f = open(outfile,"a")
        f.write(out_str + "\n")
        f.close()
    
    for line in lines:
        record = line
        line = line.strip().upper()
        line = line.strip()

        if (not line.startswith('REPORT DATE') and not line.startswith('*')
            and not line.startswith('DISTRICT') and not line.startswith('CASE NAME') 
            and not line.startswith('TOTAL CASES') and not line.startswith('END OF REPORT')
            and not line.startswith('REFERENCE NO')  
            and not line.startswith(escape) and not line.startswith(carriage) and len(line) != 0):

            countLines = countLines + 1

            if countLines == 1  :
                localOffice = record[14:17].strip() + '\t'
                unit = record[38:43].strip() + '\t'
                worker = record[58:63].strip() + '\t'

            if countLines == 2  :
                caseName = record[0:32].strip() + '\t'
                caseNumber = record[32:42].strip() + '\t'
                exceptionReason = record[50:].strip()
#                if '/PAYEE' in caseName :
#                    caseName = caseName.replace('/PAYEE','')
#                    payee = 'TRUE\t'
#                else :
#                    payee = '\t'

                out_str = (localOffice + unit + worker + caseName + caseNumber + exceptionReason)

    # write to results file                                
                if out_str != ' ':
                    f = open(outfile,"a")
                    f.write(out_str + "\n")
                    f.close()
                    countWrites = countWrites +1 
                    countLines = 0
                    heapPerson = ' '
              
    df = pd.read_csv(outfile,sep='\t',lineterminator='\n',header=None)
    df.to_excel(outfilexlsx,'Sheet1',index=False,header=False)   
    os.remove(outfile)            
    print("Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    messagebox.showinfo("Completed - WINR4210", "Completed!  " + str(countWrites) + " records converted from " + f_name + " !")

    return

#===================================================================================================

main()

