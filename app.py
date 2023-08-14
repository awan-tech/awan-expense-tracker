import os
import sys
import pygsheets
import gspread
from datetime import datetime
from dotenv import load_dotenv
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage
from oauth2client.service_account import ServiceAccountCredentials as SAC

load_dotenv()
line_bot_api = LineBotApi(os.getenv("ChanelAccessToken"))
handler = WebhookHandler(os.getenv("ChanelSecrect"))

app = Flask(__name__)

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    m2s = event.message.text
    m3s = m2s.replace(" ", "")
    m3s = m3s.split(sep = ",")
    now = datetime.now()
    s = datetime.strftime(now,'%Y-%m-%d %H:%M:%S')
    rate = ['USD', 'HKD', 'GBP', 'AUD', 'CAD', 'SGD', 'CHF', 'JPY', 'ZAR', 'SEK', 'NZD', 'THB', 'PHP', 'IDR', 'EUR', 'KRW', 'VND', 'MYR', 'CNY', 'TWD']

    if m2s == "記帳":
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="請輸入以下格式：物品, 金額, 幣別"))

    elif m2s == "幣別":
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="幣別有以下幾種:\nUSD\nHKD\nGBP\nAUD\nCAD\nSGD\nCHF\nJPY\nZAR\nSEK\nNZD\nTHB\nPHP\nIDR\nEUR\nKRW\nVND\nMYR\nCNY\nTWD"))

    elif len(m3s) < 3:
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="未依照記帳格式輸入"))

    # elif m3s[1] != int:
    #     line_bot_api.reply_message(event.reply_token,TextSendMessage(text="金額請輸入數字"))
    
    elif m3s[2] not in rate:
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="未輸入正確幣別"))

    # Record msg to google excel
    else:
        line_bot_api.reply_message(event.reply_token,TextSendMessage(text="紀錄成功"))
        pass
        #GDriveJSON就輸入下載下來Json檔名稱(金鑰, 要在GCP上申請)
        #GSpreadSheet是google試算表名稱
        GDriveJSON = 'essential-graph-394222-627ec61c8d31.json'
        GSpreadSheet = 'Test'
        GsheetKey = '1uxFCIceY1U150GRHI4nPA2JaDc85BSdojOHUVHOVJuo'
        sheet_url='https://docs.google.com/spreadsheets/d/1uxFCIceY1U150GRHI4nPA2JaDc85BSdojOHUVHOVJuo/'
            
    try:
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = SAC.from_json_keyfile_name(GDriveJSON, scope)
        gc = gspread.authorize(credentials)
        sheet = gc.open_by_url(sheet_url)
        Sheets = sheet.sheet1
    except Exception as ex:
        print('無法連線Google試算表', ex)
        sys.exit(1)
    
    
    
    
    
    
    data = [s] # data為最後要加到google sheet上的資料
    data.extend(m3s)
    Sheets.append_row(data)
    print('write success')


if __name__ == "__main__":
    app.run(port=3500)