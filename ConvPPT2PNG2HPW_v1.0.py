'''
PPT -> PNG -> HWP 파일 변경 프로그램
입력 : 변환할 Powerpoint 파일 선택
출력 : 선택한 Powerpoint 파일명과 동일한 HWP 파일 생성

버전관리
v0.9(2022.01.11, 신재호) : .이 포함된 파일명 오류 제거 , 슬라이드번호 정렬 오류 해결 
v1.0(2022.01.11, 신재호) : A3 이미지 입력시 HWP 사이즈 변경
V1.1(2022.08.17, 신재호) : PNG 사이즈를 A4규격에 맞게
'''

import os
import sys
import time
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
import shutil
import logging
import traceback

#import win32com.client
# https://github.com/mhammond/pywin32/releases/tag/b300 참고하여 버전에 맞게 다운로드
import win32com.client as win32
import win32com

#import PIL
#!pip install image
from PIL import Image

#!pip install comtypes
from comtypes.client import Constants, CreateObject

PNG_FILE_TYPE = 18
JPG_FILE_TYPE = 16

powerpoint = win32com.client.Dispatch('PowerPoint.Application')
def printhelp():
    print("\n\n################# 주의사항 ################")
    print("1. 폴더, 파일명에 . 을 넣지 않기")
    print("2. 파워포인트 인증이 안되었으면 프로그램 동작 전 미리 프로그램 한개 열기")
    print("3. 중간에 오류 발생하면 HWP 파일에 PNG 파일 넣어서 수동 제작 ")
    print("##########################################\n\n\n\n")

def resource_path(relative_path):
    '''
    PyInstaller에 의해 임시폴더에서 실행될 경우 임시폴더로 접근하는 함수
    '''
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
def getfiles():
    '''
    파일 선택하기 위한 함수
    return : 선택한 파일 디렉터리 리스트
    '''
    root = Tk()
    filelist = askopenfilenames()
    root.destroy()
    return filelist

def pathchange(path):
    '''
    디렉터리 주소 형태 변환
    '''
    npath = path.replace('/', "\\")
    return npath

def ppt2png(pptFileDir, pngfolderDir,filetype):
    '''
    PowerPoint 에서 PNG 파일 저장을 위한 함수
    param :
        pptFileDir = 변환할 ppt 파일 디렉터리 주소
        pngfolderDir = png 저장할 폴더 디렉터리 주소
        Filetype = 변환할 파일 타입 (PNG:18, JPG:16)
    참고 : https://github.com/tss12/ppt2png/blob/master/ppt2png.py
    '''

    try:
        print("   3.1. PPT TO PNG")
        time.sleep(1)
        powerpoint.Visible = True
        
        # 라이센스 오류 방지를 위해 파워포인트 창 하나 더 오픈
        startPowerPoint = win32com.client.Dispatch('PowerPoint.Application').Presentations.Open(pptFile, WithWindow=False,)
        
        #파워포인트 열어서 저장 후 종료
        ppt = powerpoint.Presentations.Open(pptFileDir, WithWindow=False,)
        ppt.SaveAs(pngfolderDir, filetype)  # 17 jpg
        ppt.Close()
        
        startPowerPoint.close()
        #powerpoint.Quit()

        print("     저장 완료 : ",pngfolderDir)
        print("     변환 완료 : ",pptFileDir)

    except Exception as e:
        # logging.error(traceback.format_exc())
        print("     PPT2PNG 오류 발생")
        print("오류명:", e)
        print("powerpoint 제품 인증 실패를 확인하세요 -> 파워포인트를 하나 더 여시오 ")
        printhelp()

     

def HwpUnitToMili(hwpunit):
    return hwpunit * 283

def PngReshape(filepath, type='A4'):
    # https://ponyozzang.tistory.com/600 참고
    A3_size = (1584,1122)
    A4_size = (790,1110)

    img = Image.open(filepath)
    if type =='A4':
        img_resize = img.resize(A4_size,Image.LANCZOS)
    else:
        img_resize = img.resize(A3_size,Image.LANCZOS)
    img_resize.save(filepath)    



def PngToHwp(pngpath, hwppath, hwpFileName):
    '''
    PNG 파일에서 HWP파일로 삽입하기 위한 부분
    param:
        pngpath: 삽입할 png 파일이 있는 폴더 디렉터리
        hwppath: 작업할 hwp 파일 위치
        hwpfilename: hwp 파일명
    
    참고:
        @ 보안 모듈 안나오게 하는 방법
        https://martinii.fun/67 참고
        1. 보안모듈 다운 링크 => https://www.hancom.com/board/devdataView.do?board_seq=47&artcl_seq=4085&pageInfo.page=&search_text=
        2. 해당 파일을 다운받고 압축을 풉니다. (JPG파일이 레지스트리 주소, dll파일이 해당 모듈입니다. 하위폴더에는 소스가 들어있습니다.)
        3. 대운받은 폴더를 한글 프로그램 폴더에 삽입 (C:\Program Files (x86)\Hnc)
        4. 윈도우+R ->   regedit  -> HKEY_CURRENT_USER\SOFTWARE\HNC\HwpAutomation\Modules 위치로 이동
        5. 우측 빈 공간에 마우스 우클릭 후 "새로 만들기(N)" - "문자열 값(S)"을 추가합니다.(파일명 : FilePathCheckerModule)
        6. 수정하기로 "값 데이터"에 dll파일을 위치를 입력합니다. (HNC 폴더에 넣었던 파일, dll 파일명 포함해서)
        7. 아래 registerModule 코드 포함하여 동작

        @ 문서 사이즈 변경
        https://martinii.fun/66
    '''
    try:
        print("   3.2. PNG TO HWP")
        #hwp = win32.Dispatch("HWPFrame.HwpObject")
        time.sleep(1)
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        # 보안모듈 적용(파일 열고닫을 때 팝업이 안나타남)
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckModule")  

        hwp.XHwpWindows.Item(0).Visible = True
        
        # HWP 페이지 양식 변경
        act = hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("ApplyTo", 3)

        item_set = pset.CreateItemSet("PageDef", "PageDef")

        item_set.SetItem("TopMargin", 0)
        item_set.SetItem("BottomMargin", 0)
        item_set.SetItem("LeftMargin", 0)
        item_set.SetItem("RightMargin", 0)
        item_set.SetItem("HeaderLen", 0)
        item_set.SetItem("FooterLen", 0)
        item_set.SetItem("GutterLen", 0)

        #A3 이미지 크기이면 HWP 사이즈 변경
        A3imageCheck = Image.open(os.path.join(pngpath, os.listdir(pngpath)[0]))
        A3imageCheckBool =False
        if A3imageCheck.width > A3imageCheck.height:
            item_set.SetItem("PaperWidth", HwpUnitToMili(420))
            item_set.SetItem("PaperHeight", HwpUnitToMili(297))
            A3imageCheckBool = True
            print("     A3 양식 변경 ")
        else:
            #A4 사이즈로 양식 변경
            item_set.SetItem("PaperWidth", HwpUnitToMili(210))
            item_set.SetItem("PaperHeight", HwpUnitToMili(297))

        act.Execute(pset)



        #슬라이드 번호로 정렬
        png_filelist = os.listdir(pngpath)
        slide_num_list = []
        for slide in png_filelist:
            slide_num_list.append(int(slide[4:-4]))
        slide_num_list.sort()


        #슬라이드 반복하여 이미지로 삽입
        for slide_num in slide_num_list:

            filename = "슬라이드"+str(slide_num)+'.PNG'
            filepath = os.path.join(pngpath, filename)

            if os.path.splitext(filepath)[1] == '.PNG':

                image = Image.open(filepath)
                
                filepath=filepath.replace("/","\\")

    
                #https://www.hancom.com/board/devmanualList.do?artcl_seq=3978 참고
                if A3imageCheckBool == True:
                    # 이미지 사이즈 변경 후
                    PngReshape(filepath,'A3')
                    # 이미지 삽입
                    hwp.InsertPicture(filepath, True, 1,(420,297))
                else:
                    # 이미지 사이즈 변경 후
                    PngReshape(filepath,'A4')
                    # 이미지 삽입
                    hwp.InsertPicture(filepath, True, 1,(297,210))
            else: 
                print(filename+"이미지 파일이 아닙니다.")

        hwp.SaveAs(hwppath+'\\'+hwpFileName)
        print("     ",hwppath, "- 저장 완료")
        hwp.Quit()
    except Exception as e:
        # logging.error(traceback.format_exc())
        print("     PNG2HWP 오류 발생")
        print("오류명: " , e)

        hwp.Quit()



if __name__ == "__main__":
    
    print("PPT 2 PNG 2 HWP 시작")
    printhelp()

    #1. 파일 선택
    print("1. 파일을 선택하세요")
    inpuit_filein=getfiles()

    #2. 변환 파일 확인
    print('2. 변환 선택한 파일 리스트')
    for num,filename_with_type in enumerate(inpuit_filein):
        print("   ",num," . ",filename_with_type)
    print("\n\n")

    #2.1 파일 이름 변경 (중간에 . 들어간 파일)
    filein = []
    for num,filename_with_type in enumerate(inpuit_filein):
        filename = filename_with_type[:-5]
        
        #점이 하나라도 있으면 변경하여 리스트에 추가
        if filename.find('.') != -1:
            #new_filename = filename.replace('.','_')
            split_temp = filename.split('/') # / 기준으로 쪼개기
            file_temp =split_temp.pop() # 마지막 위치는 파일명
            split_temp.append(file_temp.replace('.','_')) #파일명에 . 있으면 대치
            new_filename = '/'.join(split_temp) #/ 기준으로 다시 합치기
            print(new_filename)

            os.rename(filename+".pptx",new_filename+".pptx")
            print("  파일명 변환: ",filename ,"   ->   ",new_filename)

            filein.append(new_filename+'.pptx')
        #점이 없으면 리스트에 추가
        else:
            filein.append(filename+".pptx")

    #선택한 파일만큼 반복하여 동작 
    #PPT -> PNG 변경
    #PNG -> HWP 삽입
    for num,file in enumerate(filein):
        time.sleep(1)
        
        # 파일명 설정
        pptFile = file  # 피피티 파일 경로+파일명
        pptPath = os.path.split(pptFile)[0]  # 피피티 파일 경로
        pptFileName = os.path.split(pptFile)[1]  # 피피티 파일명
        pngDir=os.path.join(pptPath,os.path.splitext(pptFileName)[0])        #PNG 디렉터리 경로 (폴더명으로)

        #한글 변환후 파일명 설정
        hwpFileName = os.path.splitext(pptFileName)[0]+'.hwp'  # 한글 파일 파일명
        hwpFilePath = pngDir
        hwpFilePath = pathchange(hwpFilePath)
        pptFile = pathchange(pptFile)                     # / -> \\ 변경
        pngDir = pathchange(pngDir)

               
        print("\n\n###############",num+1,'/',len(filein),"################")
        print("ppt 파일",pptFile)
        print("png 폴더 : ", pngDir)

        ## POWERPOINT TO PNG
        ### 기존 PNG 폴더 제거
        if os.path.isdir(pngDir):  # 폴더 제거
            shutil.rmtree(pngDir)
            print("--기존 png폴더 제거\n")

        ## PPT TO PNG 동작
        ppt2png(pptFile, pngDir, PNG_FILE_TYPE)

        ## PNG TO HWP 동작 
        os.path.abspath(pngDir)
        os.path.abspath(hwpFilePath)
        PngToHwp(os.path.abspath(pngDir), os.path.abspath(hwpFilePath), hwpFileName)

    
    #파워포인트 전체 닫고 종료
    
    powerpoint.Quit()
    os.system('pause')


