import win32com.client
import datetime
import time
import sys
import os
import readchar

args = sys.argv
path = os.getcwd()

def exit_proc():
    key = readchar.readkey()
    if not key:
        sys.exit()

def sendmail():
    # Outlookのオブジェクト設定
    outlook = win32com.client.Dispatch("Outlook.Application")
    mymail = outlook.CreateItem(0)

    # 日付の取得
    vdate = datetime.date.today() # 今日の日付
    vdatefc = vdate.strftime("%Y/%m/%d") # 変換後
    dayoftheweek = ["月","火","水","木","金","土","日"]
    yobi = dayoftheweek[vdate.weekday()]

    #メールアドレスの設定
    with open(confpath) as f:
        address = [s.strip() for s in f.readlines()]
    mymail.BodyFormat = 1               #テキストタイプ
    if 0 <= len(address) <= 1:
        print("宛先が未設定です。いずれかのキーを押すと終了します")
        exit_proc()
    elif len(address) == 2:
        mymail.To = address[1]          # To
    elif len(address) == 3:
        mymail.To = address[1]          # To
        mymail.cc = address[2]          # Cc
    elif len(address) == 4:
        mymail.To = address[1]          # To
        mymail.cc = address[2]          # Cc
        mymail.Bcc = address[3]         # Bcc

    mymail.Subject = "【勤怠】 " + str(vdatefc) + " 分の送付"
    mymail.Body = "〇〇さん" + "\n\n" + "お疲れ様です。□□です。" + "\n\n" + str(vdatefc) + "(" + yobi + ")" + "の勤怠を記載しましたので、\n" + "送付致します。" + "\n\n" + "以上、よろしくお願い致します。"

    # ファイルが添付されている場合
    if len(filepath) > 0:
        mymail.Attachments.Add (filepath)

    # メール送信
    if address[1] == 0:
        mymail.Send()          # 画面非表示で送信
    else:
        mymail.Display(True)    # 作成画面を表示

if len(args) == 1:
    print("設定ファイル、添付ファイルが指定されていません。いずれかのキーを押すと終了します")
    exit_proc()
elif len(args) == 2:
    print("添付ファイル名が指定されていません。")
    print("そのまま送信する場合は y、終了する場合は y 以外を入力してください")
    key = readchar.readkey()
    if not key == 'y':
        sys.exit()
    else:
        confname = sys.argv[1]
        confpath = path + '//' + confname
        filepath = ''
        sendmail()
elif len(args) == 3:
    confname = sys.argv[1]
    confpath = path + '//' + confname
    filename = sys.argv[2]
    filepath = path + '//' + filename
    sendmail()