import win32com.client as win32
import os
import time
import shutil
from tqdm import tqdm
from tqdm import trange




# 파일 이름, 시트이름 등등을 입력받음
file_name = input("파일 이름을 입력하세요 : ")
Sheets_name = input("시트 이름을 입력하세요 : ")
Building_Num = input("동의 개수를 입력하세요 : ")
Building = [0 for z in range(int(Building_Num))]
for n in range(0,int(Building_Num),1):
    Building[n] = input("동의 이름을 입력하세요 : ")




# 실행파일이 있는 곳에 있는 엑셀 양식 파일, 엑셀 시트 열기
Path = os.getcwd()
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Add(Path+"\\"+file_name+".xlsx")
ws = wb.Sheets(Sheets_name)



# 사진 삽입 위치 초기값 및 초기화
location1=3
location2=13
location3=23
location="B"
locate1=0
locate2=0
locate3=0


# 동의 개수만큼 반복
for m in range(0,int(Building_Num),1):
    


    # 각 동의 폴더에서 .jpg로 끝나는 파일 리스트 만들기
    Path = os.getcwd()
    file_list = os.listdir(Path + "\\"+ str(Building[m]))
    file_list_jpg = [file for file in file_list if file.endswith(".jpg") or file.endswith("JPG")]   


    Image_cycle=1   # 각 동의 폴더에서 넣을 사진의 순서
    Image_Num = len(file_list_jpg)  # .jpg로 끝나는 사진들의 리스트 개수




    NUM=Image_Num // 3 # 사진이 무조곡 다 들어가는 페이지수
    NUM_1=3-(Image_Num % 3) # 마지막 페이지에서 모자란 사진 숫자
    if(Image_Num % 3 > 0):
        NUM=NUM+1 # 다들어가지 않고 남는 사진이 있으면 남는 사진이 들어간는 페이지 수 추가
        
        
        #사진 갯수가 3의 배수가 아니면 마지막 사진을 그다음번호로 복사함
        for k in range(1,NUM_1+1,1):
            shutil.copyfile(Path+"\\" + str(Building[m]) + "\\" +str(Image_Num)+".jpg", Path+"\\"+str(Building[m]) + "\\" +str(Image_Num+k)+".jpg") 


        # 사진 넣기
        for i in trange(0,int(NUM),1):
            time.sleep(0.1)
            if(i==0):
                try:
                    locate1=location+str(location1)
                    rng = ws.Range(locate1) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1


                
                    locate2=location+str(location2)
                    rng = ws.Range(locate2) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1



                    locate3=location+str(location3)
                    rng = ws.Range(locate3) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27) 
                    Image_cycle=Image_cycle+1
                except:
                    print(str(Building[m])+" "+str(Image_cycle)+".jpg 사진이 없습니다.")
                    break

            else:
                
                try:
                    locate1=location+str(location1)
                    rng = ws.Range(locate1) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)                 
                    Image_cycle=Image_cycle+1



       
                    locate2=location+str(location2)
                    rng = ws.Range(locate2) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1



                    locate3=location+str(location3)
                    rng = ws.Range(locate3) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1
                except:
                    print(str(Building[m])+" "+str(Image_cycle)+".jpg 사진이 없습니다.")
                    break
            location1=location1+32
            location2=location2+32
            location3=location3+32




    else:
        for i in trange(0,int(NUM),1):
            time.sleep(0.1)
            if(i==0):
                try:
                    locate1=location+str(location1)
                    rng = ws.Range(locate1) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1



                    locate2=location+str(location2)
                    rng = ws.Range(locate2) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1


                    locate3=location+str(location3)
                    rng = ws.Range(locate3) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1
                except:
                    print(str(Building[m])+" "+str(Image_cycle)+".jpg 사진이 없습니다.")
                    break


            else:
                try:
                    locate1=location+str(location1)
                    rng = ws.Range(locate1) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1

         
                    locate2=location+str(location2)
                    rng = ws.Range(locate2) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1


          
                    locate3=location+str(location3)
                    rng = ws.Range(locate3) 
                    Image_Path = Path+"\\" + str(Building[m]) + "\\" +str(Image_cycle) +".jpg" 
                    image = ws.Shapes.AddPicture(Image_Path, False,True, rng.Left, rng.Top, 247.68, 184.28) #조사표 사진 크기 -> 완료(2021.04.27)
                    Image_cycle=Image_cycle+1
                except:
                    print(str(Building[m])+" "+str(Image_cycle)+".jpg 사진이 없습니다.")
                    break
        
            location1=location1+32
            location2=location2+32
            location3=location3+32
    print(Building[m]+" 사진 삽입 완료.")


print("모든 사진 삽입 완료.")
input("종료하시려면 Enter 키를 눌러주십시오.")
excel.Visible=True
