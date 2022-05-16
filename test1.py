from datetime import datetime
from apscheduler.schedulers.blocking import BlockingScheduler
import datetime
import time
import os
import re
import pygsheets
import random
import pytchat
from obswebsocket import obsws, requests 

#youtube_url_re="https://www.youtube.com/watch?v=dzt9qksHoZ0"

youtube_url_re="https://www.youtube.com/watch?v=vhNjqThrIu0"

regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
gc = pygsheets.authorize(service_file= os.path.join(os.getcwd(),'read-message-343104-57d14206ec00.json'))
sht = gc.open_by_url('https://docs.google.com/spreadsheets/d/1nl5ffdH1NY-BD3R-DPlKptUPhZFeVqpZwDBujCaBS-s/')
wks = sht.worksheet_by_title("message")
wks1 = sht.worksheet_by_title("question")
wks2 = sht.worksheet_by_title("answer")

host = "localhost"
port = 4444
password = "12345678"
ws = obsws(host, port, password)
ws.connect()


def job_message():  # 每分鐘抓一次，message  

    row_count = len( wks.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix'))
    row_count_random = random.randint(1, row_count) # 隨機讀取一條
    print("第幾條:", row_count_random )   
    get_message = wks.get_value( "A"+ str(row_count_random) )  #取出內容
   # 顯示框  輸入文字
   # get_message="今天天塹真好"
    ws.call(requests.SetTextGDIPlusProperties(source="靜思語內容",text= get_message ))   
    print("靜思語內容:", get_message)   #print("每十秒傳送一段靜思語")   
    
    ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="靜思語內容",visible=True))
    time.sleep(30)
    ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="靜思語內容",visible=False))
    
def job_everyhour():
    print(" 每隔一小時，出一題目")
 #挑選題目
    row_count1 = len( wks1.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix'))
    row_count1_random = random.randint(2, row_count1) # 隨機讀取一條 第二條開始
    print("挑選到第幾行 sheet:", row_count1_random )  

    question = wks1.cell((row_count1_random,1)).value  # 取出題目
    question_image = wks1.cell((row_count1_random,2)).value  # 取出題目圖片
    answer = wks1.cell((row_count1_random,3)).value  # 取出答案   
    
 #紀錄中獎者要填入第幾行 
    row_count2 = len( wks2.get_all_values(include_tailing_empty_rows=False, include_tailing_empty=False, returnas='matrix'))+1
    print("紀錄中獎者要填入第幾行:",row_count2) 

  #  get_message1 = wks1.get_value( "A"+ str(row_count1_random) )  #取出內容
    ###
  #  print("秀到OBS上:", get_message1) 
    ###    

 #參數
    guess_name=""    #中獎人身分
    first_time="0"  #找出第一個答覆者    
    time_1 = time.time()  #紀錄現在時間，

    chat = pytchat.create(video_id= youtube_url_re ,interruptable=False)    

 #把投影片放映與文字關閉
    ws.call(requests.SetTextGDIPlusProperties(source="活動說明",text= question ))  #question 題目
    ws.call(requests.SetTextGDIPlusProperties(source="猜中者",text= "猜中者:" ))
    ws.call(requests.SetTextGDIPlusProperties(source="猜中者資訊",text= "E- Mail:" ))
    
    show_image_name = os.path.join(os.getcwd(), "ques-ans" , question_image )   #圖片
    ws.call(requests.SetSourceSettings(sourceName="活動圖片",sourceSettings={'file': show_image_name }))    
    #只要圖案打開，放上層即可
    ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="活動圖片",visible=True))
    ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="活動群組",visible=True))

    
    while time.time()- time_1 < 120:  #在幾秒內回答，不然消失 
        calc_time=120 - int( time.time()- time_1)
        print("time:", calc_time )
        if chat.is_alive():
            for c in chat.get().sync_items():
            #    print("..........:", c.datetime, c.author.name, c.message) 
                if answer == c.message and first_time=="0":
                    ws.call(requests.SetTextGDIPlusProperties(source="猜中者",text= "猜中者:"+ c.author.name ))
            #        print("呈現中獎人資訊 (網頁)")
            #        print("呈現中獎人填入電子郵件")
                    guess_name= c.author.name
                    first_time="1"
                    print("guess_name:",guess_name)
                    
                elif c.author.name == guess_name and first_time=="1":
                    if re.fullmatch(regex, c.message):  #
                        ws.call(requests.SetTextGDIPlusProperties(source="猜中者資訊",text= "E- Mail:"+ c.message ))

                        print("猜中者 c.author.name:", c.author.name)
            #            print("信箱資料 c.message:",c.message) 
                        #將資料寫入answer檔案內 
                        wks2.update_value((row_count2,1),c.datetime)
                        wks2.update_value((row_count2,2),question)
                        wks2.update_value((row_count2,3),answer)
                        wks2.update_value((row_count2,4),c.author.name)
                        wks2.update_value((row_count2,5),c.message) 
                        first_time="2"
                    else:
                        print("email is error")
   
    ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="活動圖片",visible=False))  
    ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="活動群組",visible=False))
    # ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="靜思語內容",visible=True))
    # ws.call(requests.SetSceneItemProperties(scene_name="方舟場景",item="投影片放映",visible=True))


         
'''
    開啟檔案 --- "題目":"答案"
    呈現"題目"
    中獎人身分=""
    設定第一位 = False
    
    while 十分鐘內:
        *讀取數據(中獎人，答案) 
        if 讀取答案 == "答案" and 設定第一位== False 時:
            呈現"中獎人"資訊 (網頁)
            呈現請"中獎人"填入電子郵件
            中獎人身分="中獎人"
            設定第一位= True   (取第一名，中獎人不會再被改過)
            
        if 讀取答案 == "中獎人" and 設定第一位== True 時:  
            讀取中獎人寫的內容
            if re.fullmatch(regex, email):
                呈現請"中獎人"填入電子郵件
            
                        

'''        
        
try:
    print("開始執行")
    sched = BlockingScheduler(timezone="Asia/Taipei")
    sched.add_job(job_message , 'interval', minutes=1) #每隔一分鐘送出一次
    sched.add_job(job_everyhour, 'cron', hour=15, minute=50)

    sched.start()
except:
    print("有問題")