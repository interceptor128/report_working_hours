"""
レポートをメール送信するプログラム
"""
import sys
import os
import datetime
import win32com.client
import readchar

ARGS = sys.argv
PATH = os.getcwd()


def exit_proc():
    """
    処理終了
    """
    ikey = readchar.readkey()
    if not ikey:
        sys.exit()


def sendmail():
    """
    メール送信
    """
    # Outlookのオブジェクト設定
    outlook = win32com.client.Dispatch("Outlook.Application")
    mymail = outlook.CreateItem(0)

    # 日付の取得
    vdate = datetime.date.today()  # 今日の日付
    vdatefc = vdate.strftime("%Y/%m/%d")  # 変換後
    dayoftheweek = ["月", "火", "水", "木", "金", "土", "日"]
    yobi = dayoftheweek[vdate.weekday()]

    # メールアドレスの設定
    with open(CONFPATH) as ofpath:
        address = [s.strip() for s in ofpath.readlines()]
    mymail.BodyFormat = 1               # テキストタイプ
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
    valuebody1 = "〇〇さん" + "\n\n" + "お疲れ様です。□□です。" + "\n\n"
    valuebody2 = str(vdatefc) + "(" + yobi + ")" + "の勤怠を記載しましたので、\n"
    valuebody3 = "送付致します。" + "\n\n" + "以上、よろしくお願い致します。"
    mymail.Body = valuebody1 + valuebody2 + valuebody3

    # ファイルが添付されている場合
    valuelen = len(FILEPATH)
    if valuelen > 0:
        mymail.Attachments.Add(FILEPATH)

    # メール送信
    mailtypeflg = int(address[0])
    if mailtypeflg == 0:
        mymail.Send()          # 画面非表示で送信
    else:
        mymail.Display(True)    # 作成画面を表示


if len(ARGS) == 1:
    print("設定ファイル、添付ファイルが指定されていません。いずれかのキーを押すと終了します")
    exit_proc()
elif len(ARGS) == 2:
    print("添付ファイル名が指定されていません。")
    print("そのまま送信する場合は y、終了する場合は y 以外を入力してください")
    KEY = readchar.readkey()
    if not KEY == 'y':
        sys.exit()
    else:
        CONFNAME = sys.argv[1]
        CONFPATH = PATH + '//' + CONFNAME
        FILEPATH = ''
        sendmail()
elif len(ARGS) == 3:
    CONFNAME = sys.argv[1]
    CONFPATH = PATH + '//' + CONFNAME
    FILENAME = sys.argv[2]
    FILEPATH = PATH + '//' + FILENAME
    sendmail()
