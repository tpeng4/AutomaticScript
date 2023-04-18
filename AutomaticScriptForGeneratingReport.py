import pandas as pd
import os
import matplotlib.pyplot as plt 
import sys
import docx 
from docx.shared import Cm

root_path = '.\\'
PresentMonDataDir='PresentMon'
XAxisNum = 10000
TestPlatformParameters = ['OS','CPU Type','CPU NumberOfCores','CPU NumberOfLogicalProcessors','Baseboard','SMBIOSBIOSVersion','RAM Capacity','RAM ConfiguredClockSpeed','RAM Manufacturer','GPU','DriverVersion','Config']
TestCaseParameters = ['Average GPU duration','Average Frametime','Aveage FPS','Ratio over 30 FPS','Dropped %']

## 
# this class is to store the data of each test case        
class TestCase :
    def __init__(self,case_name) :
        self.name = case_name
        self.round_number = 0
        self.round_data_tuple_list = []
        self.rounds_diagram_path = ''
        self.median_round_frametime_chart_path=''


##
# this class is to map to the structure of test report doc and store all related data. 
class ReportContent:
    def __init__(self,name) :
        self.title = name
        self.test_information = 'Test Information'
        self.test_date = ''
        self.test_platform_list=[]
        self.test_setting_path=''
        self.test_scene_path = ''
        self.test_case_list=[]
        self.data_compare_tuple_list=[]
        self.emon_data='Emon Data'
        self.Xperf_data='Xperf Data'

def parseconfigfile(file_path) :
    flag = ''
    f= open(file_path,'r',encoding='utf-8' )  
    keyword_list = ['[OS]','[CPU]','[Baseboard]','[BIOS]','[RAM]','[GPU]']
    info_list = []    
    temp_list =[]
    flag = False
    while True:    
        linedata = f.readline() 
        if flag:
            if not linedata.strip() == '' :
                temp_list.append(linedata.strip()) 

        if linedata.strip() in keyword_list: 
            temp_list =[]
            flag = True

        if linedata.strip() == ''and flag:
            flag = False
            info_list.append(temp_list)          
        if not linedata :
            break
    f.close() 
    
    platform_info_list = []
    for index, para_list in enumerate(info_list): 
        # get windows version
        if(index == 0) :
            
            str1= ''
            str2 =''
            for val in para_list : 
                if 'Name' in val :
                    str1 = val.split('=')[1] 
                if 'Build' in val:
                    str2 = val.split('=')[1]
            platform_info_list.append(str1+' '+str2) 
        # get CPU info
        if(index == 1) :
            
            str1= ''
            str2 =''
            str3=''
            for val in para_list :
                if 'Name' in val :
                    str1 = val.split('=')[1]
                if 'NumberOfCores' in val:
                    str2 = val.split('=')[1]
                if 'NumberOfLogicalProcessors' in val:
                    str3 = val.split('=')[1]
            platform_info_list.append(str1)                
            platform_info_list.append(str2)                
            platform_info_list.append(str3) 
        # get baseboard info
        if(index == 2) :
            str1= '' 
            for val in para_list :
                if 'Product' in val :
                    str1 = val.split('=')[1] 
            platform_info_list.append(str1) 
        # get bios info
        if(index == 3) :
            str1= '' 
            for val in para_list :
                if 'SMBIOSBIOSVersion' in val :
                    str1 = val.split('=')[1] 
            platform_info_list.append(str1)
        # get ram info
        if(index == 4) :
            
            str1= ''
            str2 =''
            str3=''
            for val in para_list :
                if 'Capacity' in val :
                    str1 = val.split('=')[1]
                if 'ConfiguredClockSpeed' in val:
                    str2 = val.split('=')[1]
                if 'Manufacturer' in val:
                    str3 = val.split('=')[1]
            platform_info_list.append(str1)                
            platform_info_list.append(str2)                
            platform_info_list.append(str3)
        # get GPU info
        if(index == 5) :
            
            str1= ''
            str2 =''
            for val in para_list :
                if 'Name' in val :
                    str1 = val.split('=')[1]
                if 'DriverVersion' in val:
                    str2 = val.split('=')[1]
            platform_info_list.append(str1)
            platform_info_list.append(str2)
            
    # save to doc data structure.
    total_doc_content.test_platform_list.append(platform_info_list)

def readfiles():
    files = []
    dir_list =[]
    testfiles =[]
    testcasetitle= ''
    for path in os.listdir(root_path):
    # check if current path is a file
        if os.path.isfile(os.path.join(root_path, path)):
            files.append(path) 
        else :
            dir_list.append(path)
    for dir in dir_list :
        
        dir_path = os.path.join(root_path,dir,PresentMonDataDir) 
        if(os.path.exists(dir_path)):
            for file in files:
                if dir in file :  
                    file =  file.strip('.lnk')                 
                    templist = file.split('---')
                    testcasetitle = templist[3]+' '+templist[5]+' '+templist[8]+' '+templist[6]+' '+templist[7]+' '+templist[9]+' '+templist[-1]

                    #get title name
                    total_doc_content.title = templist[3]
                    #get test date
                    total_doc_content.test_date = templist[4].split('-')[0]
                    break
            
            #append presentmon data file to file list
            for path in os.listdir(dir_path):
                # check if current path is a file
                if os.path.isfile(os.path.join(dir_path, path)):
                    testfiles.append(os.path.join(dir_path, path))     
        parsefiles(testfiles,testcasetitle)
        testfiles = []
        #parse system configuration informaion
        config_file_path = os.path.join(root_path,dir,'SystemReport-01.txt')
        if(os.path.exists(config_file_path)):
            parseconfigfile(config_file_path)
    #return testfiles


def parsefiles(files,title):
    file_list_len=len(files)
    game_title = ''
    if(file_list_len==0):
        return
    else:        
        plt.figure(dpi=200,figsize=(16,8))
        for index, file in enumerate( files ): 
         
            df=pd.read_csv(file)
            stride = int( len(df.GPUDuration)/XAxisNum)
            x = []
            GPUDuartion = []
            GPUDuartionAverage = []
            MsBetweenPresents= []
            MsBetweenPresentsAverage= []
            for i in range(XAxisNum):
                x.append(i)
                GPUDuartion.append(df.GPUDuration[i*stride])
                GPUDuartionAverage.append(round(df.GPUDuration.mean(),2))         

                MsBetweenPresents.append(df.MsBetweenPresents[i*stride])
                MsBetweenPresentsAverage.append(round(df.MsBetweenPresents.mean(),2))   
            ax=plt.subplot(2,int(file_list_len/2)+1,index+1)
            ax.plot(x,GPUDuartion,label= "GPUDuartion" )
            ax.plot(x,GPUDuartionAverage,'b' ,label= "GPUDuartionAverage")
            ax.plot(x,MsBetweenPresents,'orange' ,label= "MsBetweenPresents")
            ax.plot(x,MsBetweenPresentsAverage,'brown',label= "MsBetweenPresentsAverage" )
            ax.set_ylim(bottom=0.)
            ax.set_title(file.strip('.csv').split('\\')[-1])
            ax.set_xlabel("Frame No.")
            ax.set_ylabel("FrameTime(ms)")
            ax.legend( loc='lower right',fontsize="x-small") 
        plt.suptitle(title)
        #plt.show()
        plt.savefig(fname=title+'.png' )
        plt.close()


# generate word file
def savewordfile():
    doc = docx.Document()

# init fake data 
    #total_doc_content.title = 'Boundary'
    #total_doc_content.test_date= '2023.4.12'
   # total_doc_content.test_platform_list= [
   #     ('Microsoft Windows 11 Pro 22000.795', '12th Gen Intel(R) Core(TM) i9-12900K','16','24','ROG MAXIMUS Z690 HERO','1304','17179869184','4000','hynix','NVIDIA GeForce RTX 3090','31.0.15.1659','XeSS'),
   #     ('Microsoft Windows 11 Pro 22000.795', '12th Gen Intel(R) Core(TM) i9-12900K','8','16','ROG MAXIMUS Z690 HERO','1304','17179869184','4000','hynix','NVIDIA GeForce RTX 3090','31.0.15.1659','XeSS'),
   #     ('Microsoft Windows 11 Pro 22000.795', 'AMD Ryzen 9 5950X 16-Core Processor','16','24','ROG CROSSHAIR VIII HERO (WI-FI)','4201','8589934592','3200','G Skill Intl','NVIDIA GeForce RTX 3090','31.0.15.1659','XeSS')
   # ]
    total_doc_content.test_setting_path = 'Settings1.png'
    total_doc_content.test_scene_path = 'Boundary (64-bit Development PCD3D_SM5)  7_26_2022 2_31_36 PM.png'
    
    case1 = TestCase('Case 1')
    case1.round_number = 3 
    case1.rounds_diagram_path = 'Boundary ADL-S-i9-12900K-RL4H Ultra ADL-S-i9-12900K-RL4H 1080p DX12 8+0.png'
    case1.round_data_tuple_list = [
        ('20','30','33','80%','10%'),
        ('20','30','33','80%','10%'),
        ('20','30','33','80%','10%')
    ]
 

    case2 = case1
    case3 = case1

    total_doc_content.test_case_list.append(case1)
    total_doc_content.test_case_list.append(case2)
    total_doc_content.test_case_list.append(case3)

    doc.add_heading(total_doc_content.title)
    doc.add_heading(total_doc_content.test_information,level = 2)
    doc.add_heading('Test Data', level=3)
    doc.add_paragraph('This case was test on '+total_doc_content.test_date)
    doc.add_heading('Test Platform', level=3)
#add platform table.
    platform_table=doc.add_table(rows=len(TestPlatformParameters)+1, cols= len(total_doc_content.test_platform_list)+1,style="Table Grid")

#set first column content    
    for index, para in enumerate(TestPlatformParameters ):
        platform_table.columns[0].cells[index+1].text = para

# set other column content of platform table. 
    for index, platform in  enumerate(total_doc_content.test_platform_list) :
        platform_table.rows[0].cells[index+1].text = 'Platform'+str(index+1)
        for i,val in enumerate(platform):
            platform_table.columns[index+1].cells[i+1].text = val

# set test setting screenshot
    doc.add_heading('Test Setting', level=3)
    doc.add_picture(total_doc_content.test_setting_path, width=Cm(10))

# set test scene screenshot
    doc.add_heading('Test Scene Screenshot', level=3)
    doc.add_picture(total_doc_content.test_scene_path, width=Cm(10))

# visualize presentmon data
    doc.add_heading('Presentmon data',level = 2)
    for i,testcase in enumerate(total_doc_content.test_case_list) :        
        doc.add_heading(testcase.name,level=3)
        doc.add_heading('Data statistic',level=4)
        doc.add_paragraph('In this case, we have run %d times and conclude this data as below.'%testcase.round_number)
        testcase_table=doc.add_table(rows=testcase.round_number+1,cols=len(TestCaseParameters)+1,style="Table Grid")
        # fill out first row
        for j,para in enumerate( TestCaseParameters):
            testcase_table.rows[0].cells[j+1].text = para
        # fill out the other row
        for k, data in enumerate(testcase.round_data_tuple_list) :            
            testcase_table.rows[k+1].cells[0].text = 'Round%d'%(k+1)
            for m,value in enumerate(data):
                testcase_table.rows[k+1].cells[m+1].text =value


    doc.save('test.docx')

    total_doc_content.emon_data = 'eeee'

if __name__=="__main__":

    total_doc_content = ReportContent('')  
 

    if len(sys.argv) >1:
        if os.path.exists(sys.argv[1]) :
            root_path = sys.argv[1]
        else :
            print('Please input right path')            
    else :
        print( 'Usage: AutomaticScriptForGeneratingReport test_data_path')


    readfiles()

    savewordfile()