from concurrent.futures import ThreadPoolExecutor
import requests, json, os, progressbar
debug = False
global postNotiMsg, val, max_val, cfgData
global msgType, imgId, link, linkBtn, attCont
val = -1
max_val = 0
widgets = [
    progressbar.Percentage(), ' ',
    progressbar.SimpleProgress("(%(value_s)s/%(max_value_s)s) "),
    progressbar.Timer('[%(elapsed)s]'), ' ',
    progressbar.GranularBar(" ░▒▓█",left='[', right=']'), ' ',
    progressbar.AdaptiveETA(format_not_started='Còn lại: --:--:--', format_finished='Hoàn thành!', format="Còn lại: %(eta)8s", format_zero="Hoàn thành!"),
]
bar = progressbar.ProgressBar(max_value=max_val,redirect_stdout=True, widgets=widgets)

global req_gs, req_att, req_noti, req_mess

req_gs = "https://script.google.com/macros/s/.../exec"
req_att = "https://graph.facebook.com/v13.0/me/message_attachments?access_token=..."
req_noti = "https://script.google.com/macros/s/.../exec"
req_mess = "https://graph.facebook.com/v14.0/me/messages?access_token=..."

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def updateProgBar(label):
    global bar, val
    val += 1
    if label != "":
        print(label)
    bar.update(val)

usersList = []

def execNoti(i):
    updateProgBar(f"Gửi thông báo cho HS: {usersList[i]}")
    request_body = {
        "recipient": {
            "notification_messages_token": str(usersList[i]),
        },
        "message": attCont,
    }
    res = requests.post(req_mess, json = request_body)

import firebase_admin
from firebase_admin import firestore
from firebase_admin import credentials

def exec():
    global cfgData, req_gs, req_att, req_mess, req_noti, attCont
    cfgData = json.load(open('cfg.json', 'r', encoding='utf-8'))
    req_gs = f"https://script.google.com/macros/s/{cfgData['google_script_id']}/exec"
    req_att = f"https://graph.facebook.com/v13.0/me/message_attachments?access_token={cfgData['access_token']}"
    req_noti = f"https://script.google.com/macros/s/{cfgData['google_script_noti_id']}/exec"
    req_mess = f"https://graph.facebook.com/v14.0/me/messages?access_token={cfgData['access_token']}"
    firebase_admin.initialize_app(credentials.Certificate('key.json'))
    db = firestore.client()
    print(f"Đang lấy danh sách học sinh.")
    global max_val, val, bar, postNotiMsg, usersList, debug
    users = db.collection(u'RecurNoti').where(u'Enable', u'==', 1).stream()
    for user in users:
        usersList.append(user.to_dict()["RNToken"])
    max_val = len(usersList)
    val = -1
    bar = progressbar.ProgressBar(max_value=max_val,redirect_stdout=True, widgets=widgets)
    
    # Tạo object
    if msgType == 1:
        attCont = {
            "attachment": {
                "type": "image",
                "payload": {
                    "attachment_id": imgId
                }
            }
        }
    elif msgType == 2:
        attCont = {
            "text": postNotiMsg
        }
    elif msgType == 3:
        attCont = {
            "attachment": {
                "type": "template",
                "payload": {
                    "template_type": "button",
                    "text": postNotiMsg,
                    "buttons": [
                        {
                            "type": "web_url",
                            "url": link,
                            "title": linkBtn
                        }
                    ]
                }
            }
        }
    elif msgType == 4:
        attCont = {
            "attachment": {
                "type": "template",
                "payload": {
                    "template_type": "media",
                    "elements": [
                        {
                            "media_type": "image",
                            "attachment_id": imgId,
                            "buttons": [
                                {
                                    "type": "web_url",
                                    "url": link,
                                    "title": linkBtn
                                }
                            ]
                        }
                    ]
                }
            }
        }
    # === Phase 5: Thông tin cho Học sinh ================================ #
    cnt12 = 0
    with ThreadPoolExecutor(max_workers=int(cfgData['max_worker_count'])) as executor:
        for result in executor.map(execNoti, range(len(usersList))):
            cnt12 += 1
            print(f"{cnt12}/{len(usersList)} | Gửi thành công!")

def check():
    global msgType, postNotiMsg, imgId, link, linkBtn
    if os.path.exists('cfg.json') == False:
        print(bcolors.FAIL + "Không tìm thấy file cấu hình." + bcolors.ENDC)
        exit(0)
    
    if os.path.exists('key.json') == False:
        print(bcolors.FAIL + "Không tìm thấy file key." + bcolors.ENDC)
        exit(0)

    msgType = int(input("Loại tin nhắn [1: Chỉ Ảnh/2: Chỉ Chữ/3: Chữ và Link/4: Ảnh và Link]"))
    if msgType == 2 or msgType == 3:
        postNotiMsg = input("Tin nhắn đến học sinh: ")
        print(f"Tin nhắn: {postNotiMsg}")
        if postNotiMsg == "":
            postNotiMsg = "!"
    if msgType == 1 or msgType == 4:
        imgId = input("ID Messenger của ảnh: ")
        print(f"ID Ảnh: {imgId}")
    
    if msgType > 2:
        link = input("Link: ")
        print(f"Link: {link}")
        linkBtn = input("Chữ hiện trên nút: ")
        print(f"Chữ hiện trên nút: {linkBtn}")

check()
input("Nhấn Enter để đăng thông báo.")
exec()