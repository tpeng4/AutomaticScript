import pandas as pd
import os
import matplotlib.pyplot as plt 
import sys
import docx 
from docx.shared import Cm, RGBColor,Pt
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from datetime import datetime
from docx.enum.text import WD_ALIGN_PARAGRAPH
from tkinter import *  
from tkinter import filedialog
from tkinter import messagebox

PresentMonDataDir='PresentMon'
XAxisNum = 100
TestPlatformParameters = ['OS','CPU Type','CPU NumberOfCores','CPU NumberOfLogicalProcessors','Baseboard','SMBIOSBIOSVersion','RAM Capacity','RAM ConfiguredClockSpeed','RAM Manufacturer','GPU','DriverVersion']
TestCaseParameters = ['Average GPU duration','Average Frametime','Aveage FPS','Ratio over 30 FPS','Dropped %']
test_case_list = []
test_folder_list = []
root_path=''
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
        self.test_setting_path_list=[]
        self.test_scene_path_list = []
        self.test_case_list=[]
        self.data_compare_tuple_list=[]
        self.emon_data='Emon Data'
        self.Xperf_data='Xperf Data'
    def clear(self):
        self.title = ' '
        self.test_information = 'Test Information'
        self.test_date = ''
        self.test_platform_list=[]
        self.test_setting_path_list=[]
        self.test_scene_path_list = []
        self.test_case_list=[]
        self.data_compare_tuple_list=[]
        self.emon_data='Emon Data'
        self.Xperf_data='Xperf Data'
 

def parse_config_file(file_path) :
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
def get_platforms_list():
    lnk_files_list = sorted([ name for name in os.listdir(root_path) if name.endswith('.lnk')])  
    
    global test_case_list 
    global test_folder_list

    
    test_case_list = []
    test_folder_list = []
    
    #filter invaild test case 
    for case in lnk_files_list :
        case=case.strip('.lnk')  
        templist = case.split('---')
        temp_folder = templist[0]
        if os.path.exists(os.path.join(root_path,temp_folder,PresentMonDataDir)):
            testcasetitle = templist[3]+' '+templist[5]+' '+templist[8]+' '+templist[6]+' '+templist[7]+' '+templist[9]+' '+templist[-1]
            test_case_list.append(testcasetitle)
            test_folder_list.append(temp_folder)

def extract_case_information(selection):
    files = sorted([ name for name in os.listdir(root_path) if name.endswith('.lnk')])
    dir_list =[]
    testfiles =[]
    testcasetitle= ''
    
    dir_list=[test_folder_list[i] for i in selection] 
     
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
         
        parse_presentdata_files(testfiles,testcasetitle)
        testfiles = []
        #parse system configuration informaion
        config_file_path = os.path.join(root_path,dir,'SystemReport-01.txt')
        if(os.path.exists(config_file_path)):
            parse_config_file(config_file_path)
        
        if len(total_doc_content.test_setting_path_list)<=0 :
            dir_path = os.path.join(root_path,dir ) 
            img_path_list = sorted([os.path.join(dir_path,name) for name in os.listdir(dir_path) if name.endswith('.png')or name.endswith('.jpg')]) 
            for img in img_path_list :
            # get test setting screenshot picture file
                if 'setting' in img.lower() :
                    total_doc_content.test_setting_path_list.append(img)
            # get test scene screenshot picture file
                if 'screenshot' in img.lower():
                    total_doc_content.test_scene_path_list.append(img)
    #return testfiles


def parse_presentdata_files(files,title):
    file_list_len=len(files)
    
    testcase = TestCase(title)
    testcase.round_number=file_list_len
 
    #testcase.round_data_tuple_list
    if(file_list_len==0):
        return
    else:        
        plt.figure(dpi=200,figsize=(16,8))
        for index, file in enumerate( files ): 
         
            df=pd.read_csv(file)
            stride = int( len(df.GPUDuration)/XAxisNum)

            round_data_tuple  = []
            round_data_tuple.append("%.2f" %(df.GPUDuration.mean()))
            round_data_tuple.append("%.2f" %(df.MsBetweenPresents.mean()))
            round_data_tuple.append("%.2f" %(1000/float(df.MsBetweenPresents.mean())))
            #compute ratio of fps > 30fps
            count = len([x for x in df.MsBetweenPresents.tolist() if x <= 1000/30])
            round_data_tuple.append("%.1f%%" % (count/len(df.MsBetweenPresents) * 100))            
            round_data_tuple.append("%.2f" %(df.Dropped.sum()))
            testcase.round_data_tuple_list.append(round_data_tuple)

            # draw all round diagram for one test case 
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
        fname=title+'.png'
        plt.savefig(fname)
        plt.close()
        testcase.rounds_diagram_path=fname
    total_doc_content.test_case_list.append(testcase)

# doc heading 3 style. 
def add_heading3(doc, text) :
    table = doc.add_table(rows=1,cols=1,style='Light List Accent 5')
    cell = table.cell(0,0)    
    #shading_elm_1 = parse_xml(r'<w:shd {} w:fill="98F5FF"/>'.format(nsdecls('w')))
    #cell._tc.get_or_add_tcPr().append(shading_elm_1)
    cell.text=text
    cell.paragraphs[0].runs[0].font.size = Pt(16)   
    #cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)  
    doc.add_paragraph(" ")

# generate word file
def save_word_file():
    doc = docx.Document()

#set heading 1 style
    heading_1_style =  doc.styles['Heading 1']
    heading_1_style.font.color.rgb = RGBColor(0, 0, 0)
    heading_1_style.font.size = Pt(24)
    heading_1_style.font.name = 'Calibri Light'
    heading_1_style.paragraph_format.alignment = 1
    doc.add_heading((total_doc_content.title+' Performance testing Report').upper(),level =1)
    
    doc.add_paragraph( datetime.now().strftime('%Y-%m-%d')).alignment = WD_ALIGN_PARAGRAPH.CENTER
    #doc.add_paragraph( "DO NOT COPY OR DISTRIBUTE ")
    #doc.add_paragraph( 'Intel and Tencent Confidential ')
    paragraph = doc.add_paragraph()
    sentence = paragraph.add_run('DO NOT COPY OR DISTRIBUTE \n Intel Confidential')
    sentence.italic = True
    sentence.font.color.rgb = RGBColor(255, 0, 0)
    sentence.font.name = 'Calibri Light'
    sentence.font.size= Pt(14)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_page_break()

    doc.add_heading(total_doc_content.test_information,level = 2)


    doc.add_heading('Test Date', level=3)
    doc.add_paragraph('This case was test on '+total_doc_content.test_date)
    

    #doc.add_heading('SYSTEM CONFIGURATIONS', level=3)
    add_heading3(doc,'SYSTEM CONFIGURATIONS' )
#add platform table.
    platform_table=doc.add_table(rows=len(TestPlatformParameters)+1, cols= len(total_doc_content.test_platform_list)+1)#,style="Table Grid")
    platform_table.style = 'Medium Grid 3 Accent 1'
#set first column content    
    for index, para in enumerate(TestPlatformParameters ):
        platform_table.columns[0].cells[index+1].text = para

# set other column content of platform table. 
    for index, platform in  enumerate(total_doc_content.test_platform_list) :
        platform_table.rows[0].cells[index+1].text = 'Platform'+str(index+1)
        for i,val in enumerate(platform):
            cell=platform_table.columns[index+1].cells[i+1]
            cell.text = val
            cell.paragraphs[0].runs[0].font.size = Pt(9)   


# set test setting screenshot
    #doc.add_heading('GAME SETTING', level=3)
    doc.add_paragraph(" ")
    add_heading3(doc, 'GAME SETTING')
# set config parameter table(2*4) 
    config_table = doc.add_table(rows=4, cols=2,style = "Light Grid Accent 6")
    config_table.columns[0].cells[0].text = "QUALITY LEVLE"
    config_table.columns[0].cells[1].text = "API VERSION"
    config_table.columns[0].cells[2].text = "RESOLUTION"
    config_table.columns[0].cells[3].text = "GOAL"
    for i in range(0,4):
        config_table.columns[0].cells[i].paragraphs[0].alignment = 2 
    if len(total_doc_content.test_case_list) >0 :        
        
        temp_config_list = total_doc_content.test_case_list[0].name.split(' ')
        config_table.columns[1].cells[0].text =  temp_config_list[4]
        config_table.columns[1].cells[1].text =  temp_config_list[5]
        config_table.columns[1].cells[2].text =  temp_config_list[3]
        config_table.columns[1].cells[3].text =  temp_config_list[1]

    for test_setting_path in total_doc_content.test_setting_path_list:
        doc.add_picture(test_setting_path, width=Cm(15))


# set test scene screenshot
    doc.add_heading('GAME SCENE SCREENSHOT', level=3)
    for test_scene_path in total_doc_content.test_scene_path_list:
        doc.add_picture(test_scene_path, width=Cm(15))
# set test summary    
    add_heading3(doc, 'TESTING RESULTS')
    doc.add_paragraph('SUMMARY')



# visualize presentmon data
    doc.add_heading('Detail performance data based on PresentMon',level = 2)
    for i,testcase in enumerate(total_doc_content.test_case_list) :        
        doc.add_heading(testcase.name,level=3)
        doc.add_heading('Data statistic',level=4)
        doc.add_paragraph('In this case, we have run %d times and conclude this data as below.'%testcase.round_number)
        testcase_table=doc.add_table(rows=testcase.round_number+1,cols=len(TestCaseParameters)+1)
        testcase_table.style = 'Medium Grid 3 Accent 1'
        # fill out first row
        for j,para in enumerate( TestCaseParameters):
            testcase_table.rows[0].cells[j+1].text = para
        # fill out the other row
        for k, data in enumerate(testcase.round_data_tuple_list) :            
            testcase_table.rows[k+1].cells[0].text = 'Round%d'%(k+1)
            for m,value in enumerate(data):
                cell=testcase_table.rows[k+1].cells[m+1]
                cell.text =value
                cell.paragraphs[0].runs[0].font.size = Pt(9)  
        
        doc.add_heading('Diagram all rounds',level=4)
        doc.add_picture(testcase.rounds_diagram_path,width=Cm(15))


    filename = total_doc_content.title+'.docx'
    if(os.path.exists(filename)) :
        os.remove(filename)
    doc.save(filename)

    total_doc_content.emon_data = 'emon_data'

def button_choose_dir_fun():
    
    
    listbox_global.delete(0, END)
    global root_path
    root_path=filedialog.askdirectory() 
    label_global.config(text=root_path)
    if root_path !='' :
        get_platforms_list()

        if len(test_case_list)>0 :
            for item in test_case_list:
                listbox_global.insert(END, item)
        else:
            messagebox.showinfo("Warning", "Please choose right folder")
def button_generate_report_fun():

    selection = listbox_global.curselection() 
        
    extract_case_information(selection)

    save_word_file() 
    messagebox.showinfo("Notice", "Report was generated")
    
    listbox_global.delete(0, END)
    
    # empty all content.
    total_doc_content.clear()

def create_window():

    root = Tk()
    #root.withdraw()
 
    root.title("GenerateReport")
    root.geometry('500x400+700+200')
    button_choose_dir = Button(root, text="Choose Directory", command=button_choose_dir_fun) 
    button_choose_dir.grid() 

    label_dir= Label(root,text='Please choose test data folder')
    label_dir.grid(column=1, row=0)

    Label(root,text='Test Case List').grid(column=0,row=1)
    listbox_test_case = Listbox(root, bg="white", fg="black", bd=5, height=10, width=20, font=("Arial", 14), selectmode='multiple')
    listbox_test_case.grid(column=0, row=2)

    button_process=Button(root,text='Generate Report',command=button_generate_report_fun)
    button_process.grid(column=0,row=3)
    button_exit=Button(root,text='Exit',command=root.destroy)
    
    button_exit.grid(column=1,row=3)

    global label_global 
    label_global= label_dir

    global listbox_global
    listbox_global = listbox_test_case

    root.mainloop() 
if __name__=="__main__":
    
    total_doc_content = ReportContent('')  
    
    create_window()

    
 
    #if len(sys.argv) >1:
    #    if os.path.exists(sys.argv[1]) :
    #        root_path = sys.argv[1]
    #    else :
    #        print('Please input right path')            
    #else :
    #    print( 'Usage: AutomaticScriptForGeneratingReport test_data_path')


    #read_files()

    #save_word_file()