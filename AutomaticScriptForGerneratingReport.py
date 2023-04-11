import pandas as pd
import os
import matplotlib.pyplot as plt 
import sys

root_path = '.\\'
PresentMonDataDir='PresentMon'
XAxisNum = 10000
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
                    break

            for path in os.listdir(dir_path):
                # check if current path is a file
                if os.path.isfile(os.path.join(dir_path, path)):
                    testfiles.append(os.path.join(dir_path, path))
                    
        parsefiles(testfiles,testcasetitle)
        testfiles = []
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
if __name__=="__main__":
    
    if len(sys.argv) >1:
        if os.path.exists(sys.argv[1]) :
            root_path = sys.argv[1]
        else :
            print('Please input right path')
    readfiles()
