
import pandas as pd
import numpy as np
from pandas import ExcelWriter
from pandas import ExcelFile

#Import Primers Sequence
primers=pd.read_excel(r'C:\Users\thaol\Documents\Visual Studio\primers.xlsx',header=None)

#Import sample sheet matrix
data = pd.read_excel(r'C:\Users\thaol\Documents\Visual Studio\Test.xlsx','Test')

#Begin read of forward primers
f=data.columns
#empty arrays
trash=[]; #Puts empty arrays into trash array
samplename=[]; #Array of sample names
fp=[]; #Array of forward primers
rp=[]; #Array of reverse primers
fpseq=[];
rpseq=[];
#sorting for forward primers
for x in range(2,len(f)): #passes through each column
    for y in range(1,len(data[f[x]])): #passes through each row within 'x' column
        a=data[f[x]]
        if (str(a[y]) == str(np.nan)):
            trash.append(float(y))
        else:
            fp.append(f[x])
            rp.append(data.iloc[y,0]) #Appends reverse primer associated to sample
            samplename.append(a[y])

#Link forward primers with sequence        
for fpn in range (0,len(fp)):
    for pn in range (0,len(primers)):
        if (fp[fpn]==primers.iloc[pn,0]):
            fpseq.append(primers.iloc[pn,1])

#Link reverse primers with sequence                 
for rpn in range (0,len(rp)):
    for pn in range (0,len(primers)):
        if (rp[rpn]==primers.iloc[pn,0]):
            rpseq.append(primers.iloc[pn,1])

#adding headers to each column for empty dataframe
ss = pd.DataFrame(columns=['Sample ID','Description','I7_Index_ID','index','I5_Index_ID','index2'])

#appending each column with each sample info
for n in range(0,len(samplename)):
    ss=ss.append({'Sample ID':samplename[n],'I7_Index_ID':rp[n],'index':rpseq[n],'I5_Index_ID':fp[n],'index2':fpseq[n]},ignore_index=True)
    
#ss=ss.append({'Sample ID':np.transpose(samplename),'I7_Index_ID':np.transpose(rp),'index':np.transpose(rpseq),'I5_Index_ID':np.transpose(fp),'index2':np.transpose(fpseq)},ignore_index=True)
df=pd.DataFrame(ss)

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('iseq_sheet.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1', index=False)

# Close the Pandas Excel writer and output the Excel file.
writer.save()

#print(samplename)
#print(fp)
#print(fpseq)
#print(rp)
#print(rpseq)