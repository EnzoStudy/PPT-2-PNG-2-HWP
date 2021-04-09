import win32com.client as win32
import win32com
# https://github.com/mhammond/pywin32/releases/tag/b300 참고하여 버전에 맞게 다운로드

import sys
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

powerpoint = win32com.client.Dispatch('PowerPoint.Application')

def ppt2png(pptFileDir, pngfolderDir,filetype):
      ##참고 : https://github.com/tss12/ppt2png/blob/master/ppt2png.py  ##
    try:

        powerpoint.Visible = True

        ppt = powerpoint.Presentations.Open(pptFileDir)

        ppt.SaveAs(pngfolderDir, filetype)  # 17 jpg

        # ppt.Close()
        # powerpoint.Quit()
    except:
        logging.error(traceback.format_exc())
        print("PPT2PNG 오류 발생")
     


def getfiles():
    # %% 이미지파일 선택
    root = Tk()
    filelist = askopenfilenames()
    root.destroy()
    return filelist


def getdirpath():
    root = Tk()
    # root.withdraw()
    dir_path = askdirectory(parent=root, initialdir="/",
                            title='Please select a directory')
    root.destroy()
    return(dir_path)


def pathchange(path):
    npath = path.replace('/', "\\")
    return npath



if __name__ == "__main__":
    ########파일 경로 및 디렉터리 설정############

    print("파일을 선택하세요")
    filein=getfiles()
    print("1. pdf , 2. png파일  3. jpg파일")
    temp=input()

    for i in filein:
        print(i)
    print("를 변환합니다..")

    if temp== '1':
        filetype=32
    elif temp=='2':
        filetype=18
    elif temp=='3':
        filetype=17


    n=1
    for i in filein:


        pptFile = i  # 피피티 파일 경로+파일명
        pptPath = os.path.split(pptFile)[0]  # 피피티 파일 경로
        pptFileName = os.path.split(pptFile)[1]  # 피피티 파일명
        pngDir=os.path.join(pptPath,os.path.splitext(pptFileName)[0])        #PNG 디렉터리 경로 (폴더명으로)

        pptFile = pathchange(pptFile)                     # / -> \\ 변경
        pngDir = pathchange(pngDir)


        print("\n\n###############",n,'/',len(filein),"################")
        print("ppt 파일",pptFile)
        print("png 폴더 : ", pngDir)

        ################메인 함수 #####################
        ppt2png(pptFile, pngDir,filetype)  # ppt -> PNG

        print("저장 완료 : ",pngDir)
        n=n+1
    
    powerpoint.Quit()
    os.system('pause')