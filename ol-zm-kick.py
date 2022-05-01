""" outlook の今日の予定から URL を取り出して、zoom か Teams を起動する """

#import pyperclip
import webbrowser
import win32com.client
import datetime
#import win32timezone
import time

kick_url = ''
zm_tag = 'Meeting)'
zm_tag2 = 'Zoomミーティングに参加する'
tm_tag = '会議に参加するにはここをクリックしてください'

zm_key_idx = -1
tm_key_idx = -1
find_flg = 0

zoom_id = ''
zoom_pc = ''
xmi = 0
xpw = 0
xpc = 0

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

calendar = outlook.GetDefaultFolder(9).items
calendar.Sort("[Start]")
calendar.IncludeRecurrences = "True"

select_items = [] # 指定した期間内の予定を入れるリスト

# 予定を抜き出したい期間を指定
today_date = datetime.date.today() # 今日だけ
# today_date = datetime.date.fromisoformat('2022-04-21') # test 
now_time = datetime.datetime.now() # 今の時間
wait_time = now_time + datetime.timedelta(seconds=900) # 今の時間 + 15min
Min = wait_time.strftime('%H:%M:%S') #15min の加算補正あり
# Min = '10:14:00' #debug 用

# restrict appointments to specified range
calendar = calendar.Restrict("[Start] >= '" + str(today_date) +
                             "' AND [END] <= '" + str(today_date + datetime.timedelta(days=1)) + "'")

# 今日のデータを取り出し
for item in calendar:
    if today_date == item.start.date():
        select_items.append(item)
    if today_date < item.start.date():
        break 

text =""
find_flg = 0 

# 抜き出した予定の詳細を表示
for select_item in select_items:
    if select_item.body == '' or select_item.body == ' '  or select_item.subject.startswith('Canceled:') : 
        continue
     #本文が空かキャンセルなので、次の予定に飛ぶ

    print("件名：", select_item.subject)
    start_time = select_item.start.time().strftime('%H:%M:%S')
    end_time = select_item.end.time().strftime('%H:%M:%S')

    if start_time <= Min and end_time > Min :  #開始15分前～終了15分前まで
        print("該当会議あり：", select_item.subject)
        lines = select_item.body.split()

        # なぜかfindでは動かない
        try : #zoom のタグを探す
            zm_key_idx = lines.index(zm_tag)
        except ValueError :
            zm_key_idx = -1
        if zm_key_idx != -1 :
            find_flg = 1 #zoom でかつ直接 URL を叩ける
            break

        try : #zoom のタグ２を探す
            if zm_key_idx == -1 :
                zm_key_idx = lines.index(zm_tag2)
        except ValueError :
            zm_key_idx = -1
 
        try : #teams のタグを探す
            tm_key_idx = lines.index(tm_tag)
        except ValueError :
            tm_key_idx = -1

        #どちらも見つからない → 次の予定へ
        if zm_key_idx == -1 and tm_key_idx == -1 :
            print ('ERR web会議ではありません')
            continue
        else : #どちらかがある
            find_flg = 2
            break

# 結果の判定
if find_flg == 0 : #一つも見つからない場合
    print ('ERR 該当会議がありません')
    input ()
    exit()
   
if zm_key_idx != -1 and tm_key_idx != -1 :
    print ('ERR zoom と teams の２つの URL があります')
    input ()
    exit()

if find_flg == 1 : #zoom のパスコード込みだったら次の行
    kick_url = lines[zm_key_idx+1]
else : # pass-code が別にあり
    if zm_key_idx != -1 :
        try:
	        xmi = lines.index('ミーティングID:')

        except ValueError :
	        # エラーを表示
	        print ('ERR ID 取得に失敗 / ミーディングID: が見当たりません')
	        input ()
	        exit()

        zoom_id = lines[xmi+1] + lines[xmi+2] + lines[xmi+3]

        try:
	        xpc = lines.index('パスコード:')
        except ValueError :
	        pass

        if xpc == 0 :
	        try:
		        xpw = lines.index('パスワード:') #古い記載の救済
	        except ValueError :
		        pass

        if xpw != 0 : #パスワード発見
	        zoom_pc = lines[xpw+1]

        if xpc != 0 : #パスコード発見 パスワードよりパスコードを優先
	        zoom_pc = lines[xpc+1]

        if zoom_pc =='' :
	        print ('WRN Pass 取得に失敗 / パスコード: が見当たりませんが起動します')
	        input ()

        kick_url = 'zoommtg://zoom.us/join?confno=' + zoom_id + '&pwd=' + zoom_pc

if tm_key_idx != -1 : #Teams だったら次の行の前後削除 '<URL>' なので
    kick_url = lines[tm_key_idx+1]
    kick_url = kick_url[1:len(kick_url)-1]

Min2 = now_time.strftime('%H:%M:%S')  #実時間
# Min2 = '09:56:00' # debug 用
 
sleep_time = int(start_time[0:2])*3600 + int(start_time[3:5])*60 + int(start_time[6:8]) - int(Min2[0:2])*3600 -int(Min2[3:5])*60 - int(Min2[6:8]) - 30 #分と秒で 30秒前に起動
if sleep_time < 0 : #残りが30秒以下の場合、すぐ起動する
    sleep_time = 0
 
print (f'起動まで { sleep_time } 秒 待ちます' )
time.sleep(sleep_time)
 
webbrowser.open(kick_url)
#input()