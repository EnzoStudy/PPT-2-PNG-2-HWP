import win32com.client as win32 
import win32com
#https://github.com/mhammond/pywin32/releases/tag/b300 참고하여 버전에 맞게 다운로드

import os
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import askdirectory

from PIL import Image
#!pip install image
from comtypes.client import Constants, CreateObject
#!pip install comtypes
import shutil
import logging
import traceback



def filelistsort(filelist):
    nlist=list
    for i in filelist:
       #nlist.append( i[1].split('.')[0])
       print(i[1])



def ppt2png(pptFileDir,pngfolderDir):
  ##참고 : https://github.com/tss12/ppt2png/blob/master/ppt2png.py  ##

    try:
        powerpoint = win32com.client.Dispatch('PowerPoint.Application')

        powerpoint.Visible = True

        ppt = powerpoint.Presentations.Open(pptFileDir)

        ppt.SaveAs(pngfolderDir,18)  #17 jpg
    
        ppt.Close()
        powerpoint.Quit()
    except:
        logging.error(traceback.format_exc())
        print("PPT2PNG 오류 발생")
        ppt.Close()
        powerpoint.Quit()

def getfiles():
    # %% 이미지파일 선택
    root=Tk()
    filelist=askopenfilenames()
    root.destroy()
    return filelist

def getdirpath():
    root = Tk()
    # root.withdraw()
    dir_path = askdirectory(parent=root,initialdir="/",title='Please select a directory')
    root.destroy()
    print(dir_path)

    return(dir_path)


def pathchange(path):
    npath=path.replace('/',"\\")
    return npath

def PngToHwp(dirpath,hwppath):
    try:
        hwp=win32.Dispatch("HWPFrame.HwpObject")
        hwp.XHwpWindows.Item(0).Visible=True           
        #hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈 적용(파일 열고닫을 때 팝업이 안나타남)
        
        #양식 변경 
        hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)
        hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(0)
        hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(0)
        hwp.HParameterSet.HSecDef.PageDef.TopMargin = hwp.MiliToHwpUnit(0)
        hwp.HParameterSet.HSecDef.PageDef.BottomMargin = hwp.MiliToHwpUnit(0)
        hwp.HParameterSet.HSecDef.PageDef.HeaderLen = hwp.MiliToHwpUnit(0.0)
        hwp.HParameterSet.HSecDef.PageDef.FooterLen = hwp.MiliToHwpUnit(0.0)
        hwp.HParameterSet.HSecDef.PageDef.GutterLen = hwp.MiliToHwpUnit(0.0)
        hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)
        hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)  # 문서 전체 변경
        hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)

        filelist=os.listdir(dirpath)

        # for file in filelist:
        
        for i in range(1,len(filelist)+1):
            file="슬라이드"+repr(i)+'.PNG'
            #os.getcwd()
            filepath=os.path.join(os.getcwd(),file)
            if os.path.splitext(filepath)[1]=='.PNG':

                #print(file)
                image = Image.open(filepath)
                resize_image = image.resize((810 ,1140))
                image.thumbnail((810 ,1140), Image.ANTIALIAS)
                image.save(filepath)
                #filepath=filepath.replace("/","\\")
                
                print(filepath.replace('\\',"/"))
                hwp.InsertPicture(filepath,True,2,None,None,None)
            
                
            else :print(file+"이미지 파일이 아닙니다.")
                    
        hwp.SaveAs(hwppath)
        print(hwppath,"- 저장 완료")
        hwp.Quit()
    except:
        logging.error(traceback.format_exc())
        print("PNG2HWP 오류 발생")
        hwp.Quit()

if __name__ == "__main__":

    #########파일 경로 및 디렉터리 설정############
    pptFile=getfiles()[0]                          #피피티 파일 경로+파일명
    pptPath=os.path.split(pptFile)[0]               #피피티 파일 경로
    pptFileName=os.path.split(pptFile)[1]           #피피티 파일명
    hwpFileName=os.path.splitext(pptFileName)[0]+'.hwp' #한글 파일 파일명
    
    #pngDir=os.path.join(pptPath,os.path.splitext(pptFileName)[0])        #PNG 디렉터리 경로 (폴더명으로)
    pngDir=os.path.join(pptPath,"PNG폴더")         #PNG 디렉터리 경로   (PNG폴더 명으로 만들기)
    hwpFilePath=os.path.join(pptPath,hwpFileName)

    pptFile=pathchange(pptFile)                     # / -> \\ 변경
    pngDir=pathchange(pngDir)
    
    print(os.path.isdir(pngDir))
    print(pngDir)
    
    if os.path.isdir(pngDir):                   #폴더 제거
        shutil.rmtree(pngDir)
        print("png폴더 제거")

    print("png 폴더 : ",pngDir )
    print("hwp 파일 : ",hwpFilePath)
    ################메인 함수 #####################

    ppt2png(pptFile,pngDir)                         #ppt -> PNG
    

    os.path.abspath(pngDir)
    os.path.abspath(hwpFilePath)
    PngToHwp(os.path.abspath(pngDir),os.path.abspath(hwpFilePath))                         #PNG -> hwp
  

    if os.path.isdir(pngDir):                   #폴더 제거
        shutil.rmtree(pngDir)
        print("png폴더 제거")

    os.system('pause')
