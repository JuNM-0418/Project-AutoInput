import win32com.client as win32
import os
import time
import shutil
from tqdm import tqdm
from tqdm import trange

# 화살표를 복사해주는 함수
def Input_Arrow(location_Row):
  ws.Range("Z1:AF4").Copy(ws.Range("U"+str(location_Row)))

# 조사표 내용을 조사사진에 넣어주는 함수
def Input_Contents(location_Row, Contents_Cycle):
    #41,43,45,47,49    25,38
    ws.Cells(int(location_Row)+8, 15).Value = ws.Cells(int(Contents_Cycle)+9, 41)
    ws.Cells(int(location_Row)+8, 17).Value = ws.Cells(int(Contents_Cycle)+9, 43)
    ws.Cells(int(location_Row)+8, 19).Value = ws.Cells(int(Contents_Cycle)+9, 45)
    ws.Cells(int(location_Row)+8, 21).Value = ws.Cells(int(Contents_Cycle)+9, 47)
    ws.Cells(int(location_Row)+8, 23).Value = ws.Cells(int(Contents_Cycle)+9, 49)
    ws.Cells(int(location_Row)+8, 25).Value = ws.Cells(int(Contents_Cycle)+9, 38)
    Contents_Cycle = Contents_Cycle + 1
    return(Contents_Cycle)

# 설명번호와 사진번호가 맞는지 확인해주는 엑셀수식을 삽입해주는 함수
def Check_ImageNum(location_Row, location_Col):
    ws.Cells(int(location_Row)+1, 21).Value = "=IF(MID(O"+str(int(location_Row)+1)+",SEARCH(\".\",O"+str(int(location_Row)+1)+",1),3)=MID(W"+str(int(location_Row)+8)+",SEARCH(\".\",W"+str(int(location_Row)+8)+",1),3),\"\",\"번호확인\")" 
    ws.Cells(int(location_Row)+1, 21).Font.Color = -16776961
    return()


# 설명부분 내용을 합성해주는 엑셀수식을 삽입해주는 함수
def Combine_Explanation(location_Row, location_Col):
    if(ws.Cells(int(location_Row)+8, 25).Value == "균열"):
        ws.Cells(int(location_Row)+4, 15).Value = "="+ location_Col + str(int(location_Row) + 8) + "&"+"\"균열\"" 
    else:
        ws.Cells(int(location_Row)+4, 15).Value = ws.Cells(int(location_Row)+8, 25)
    return()


# 이전 행(Row)을 받고 다음 사진이 삽입될 행(Row)의 넘버를 반환해주는 함수
def Next_Location(location_Row):
    location_Row = location_Row + 32
    return(location_Row)


# 행, 열, 사진경로, 폴더이름, 사진 번호를 받고 사진을 해당위치에 삽입 및 다음 사진 넘버를 반환 해주는 함수
def Input_Image(location_Col, location_Row, Path, Building,  Image_Cycle):
    location = location_Col + str(location_Row)
    rng = ws.Range(location) 
    Image_Path = Path+"\\" + str(Building) + "\\" + str(Image_Cycle) + ".jpg" 
    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28)
    Image_Cycle = Image_Cycle + 1
    return(Image_Cycle) 


# 파일 이름, 시트이름 등등을 입력받음
file_name = input("파일 이름을 입력하세요 : ")
#Survey_Sheets_name = input("조사표 시트 이름을 입력하세요 : ")

Building_Num = input("동의 개수를 입력하세요 : ")

for Q in range(0,int(Building_Num),1):
    Sheets_name = input("조사사진 시트 이름을 입력하세요 : ") 
    Building = input("동의 이름을 입력하세요 : ")


    # 실행파일이 있는 위치에 있는 엑셀 양식 파일, 엑셀 시트 열기
    Path = os.getcwd()
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    wb = excel.Workbooks.Open(Path + "\\" + file_name + ".xlsx")
    ws = wb.Sheets(Sheets_name)
    #ws_Survey = wb.Sheets(Survey_Sheets_name)


    location_Row_1 = 3 # 페이지 내 첫번째 시작 행
    location_Row_2 = 13 # 페이지 내 두번쨰 시작 행
    location_Row_3 = 23 # 페이지 내 세번째 시작 행
    location_Col_1 = "B" # 고정 열
    location_Col_2 = "U" # 번호확인 함수 위치
    location_Col_3 = "O" # 설명확인 함수 위치
    Image_Cycle = 1   # 각 동의 폴더에서 넣을 사진의 순서
    Contents_Cycle =  1 # 조사표 내용이 삽입되는 순서


        

    # 각 동의 폴더에서 .jpg로 끝나는 파일 리스트 만들기
    file_list = os.listdir(Path + "\\" + str(Building))
    file_list_jpg = [file for file in file_list if file.endswith(".jpg") or file.endswith("JPG")]   


    Image_Cycle = 1 # 사진번호 초기화
    Image_Num = len(file_list_jpg)  # .jpg로 끝나는 사진의 개수




    NUM=Image_Num // 3 # 사진이 무조건 다 들어가는 페이지수
    NUM_1 = 3 - (Image_Num % 3) # 마지막 페이지에서 모자란 사진 숫자
    if(Image_Num % 3 > 0):
        NUM = NUM +1 # 다들어가지 않고 남는 사진이 있으면 남는 사진이 들어간는 페이지 수 추가
        
        
        #사진 갯수가 3의 배수가 아니면 마지막 사진을 그다음번호로 복사함 
        for k in range(1, NUM_1+1, 1):
            shutil.copyfile(Path + "\\" + str(Building) + "\\" + str(Image_Num) + ".jpg", Path + "\\" + str(Building) + "\\" + str(Image_Num+k) + ".jpg") 


    # 사진 넣기
    for i in trange(0, int(NUM), 1):
        time.sleep(0.1)
        try:
            # 사진, 화살표 삽입 및 조사표 내용 삽입
            Image_Cycle = Input_Image(location_Col_1, location_Row_1, Path, Building, Image_Cycle)
            Contents_Cycle = Input_Contents(location_Row_1, Contents_Cycle)
            Input_Arrow(location_Row_1)
            Image_Cycle = Input_Image(location_Col_1, location_Row_2, Path, Building,  Image_Cycle)
            Contents_Cycle = Input_Contents(location_Row_2, Contents_Cycle)
            Input_Arrow(location_Row_2)
            Image_Cycle = Input_Image(location_Col_1, location_Row_3, Path, Building, Image_Cycle)
            Contents_Cycle = Input_Contents(location_Row_3, Contents_Cycle)
            Input_Arrow(location_Row_3)


        except:
            print(str(Building) + " " + str(Image_Cycle) + ".jpg 사진이 없습니다.")
            break
        # 엑셀 수식 삽입
        Combine_Explanation(location_Row_1,  location_Col_2)
        Combine_Explanation(location_Row_2,  location_Col_2)
        Combine_Explanation(location_Row_3,  location_Col_2)
        Check_ImageNum(location_Row_1, location_Col_2)
        Check_ImageNum(location_Row_2, location_Col_2)
        Check_ImageNum(location_Row_3, location_Col_2)


        location_Row_1 = Next_Location(location_Row_1)
        location_Row_2 = Next_Location(location_Row_2)
        location_Row_3 = Next_Location(location_Row_3)




    print(Building + " 사진 삽입 완료.")
    #wb.Save()
    #excel.Quit()

#Display_AnothType()
print("모든 사진 삽입 완료.")
input("종료하시려면 Enter 키를 눌러주십시오.")
excel.Visible=True