from concurrent.futures import ThreadPoolExecutor
import random
from shutil import copy2
from datetime import datetime
import time
import requests, json, imgkit, os, openpyxl, progressbar
debug = False
global wb, sheet, postNotiMsg, postNoti, today, val, max_val, cfgData
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

def execClass(i):
    # === Phase 1: Phân tích file ================================ #
    copy2("imgGen/def.html",f"imgGen/index{i}.html")
    className = str(sheet.cell(3,i+1).value)
    updateProgBar("")
    dataArray = [[]]
    tmp = str()
    for j in range(1, 31):
        if (j-1) % 5 == 0:
            tmp += "Thứ " + str((int)((j-1)/5+2)) + ": "
            dataArray.append([])
        tmp2 = str(sheet.cell(j+3,i+1).value).split('-')[0].removesuffix(' ')
        if tmp2 == "None":
            tmp2 = "⌀"
            dataArray[(int)((j-1)/5)].append(' ')
        else:
            dataArray[(int)((j-1)/5)].append(tmp2)
        tmp += tmp2
        if j%5 != 0:
            tmp += ", "
        else:
            if j != 30:
                tmp += "\n"
    
    # === Phase 2: Xuất file ảnh ================================ #
    updateProgBar("")
    with open(f'imgGen/index{i}.html', 'r', encoding='utf-8') as file:
        filedata = file.read()
    filedata = filedata.replace('ClassName', className)
    filedata = filedata.replace('ChangeDate', str(today))
    for i2 in range(0,5):
        for j2 in range(0,6):
            rep = 'R'+str(int((i2+1)))+'C'+str(int((j2+1)))
            filedata = filedata.replace(rep, dataArray[j2][i2])
    with open(f'imgGen/index{i}.html', 'w', encoding='utf-8') as file:
        file.write(filedata)
    updateProgBar("")
    options = {'enable-local-file-access': None, 'width': 1200, 'disable-smart-width': '', "quiet": ""}
    imgRes = "output/" + className + ".jpg"
    imgkit.from_file(f"imgGen/index{i}.html", imgRes, options=options)

    # === Phase 3: Đăng ảnh lên API Messenger ================================ #
    updateProgBar(f"{className} | Đang đăng ảnh lên Messenger")
    payload = {
        'message': (None, {"attachment": {"type": "image", "payload": {"is_reusable":'true'}}}),
    }
    files = {
        'filedata': ('test', open(imgRes, 'rb'), 'image/png')
    }
    retries = 0
    while retries<5:
        try:
            res = requests.post(req_att, data=payload, files=files)
            attID = json.loads(res.content)
            if res.status_code == 200:
                break
        except:
            retries += 1
            print(f"{className} | Đang đăng ảnh lên API Messenger (Lần {retries}/5)")
    return [className, attID["attachment_id"], tmp, str(today)]

usersList = []

def execNoti(i):
    updateProgBar(f"Gửi thông báo cho HS: {usersList[i]}")
    txt = (str(random.choice(cfgData['noti_post_msg_first']))).replace("<time>", str(datetime.now().strftime("%H:%M %d/%m/%Y"))) + " " + random.choice(cfgData['noti_post_msg_second']) + postNotiMsg
    request_body = {
        "recipient": {
            "notification_messages_token": str(usersList[i]),
        },
        "message": {
            "attachment": {
                "type": "template",
                "payload": {
                    "template_type": "button",
                    "text": txt,
                    "buttons": [{
                        "type": "postback",
                        "title": "XEM NGAY!",
                        "payload": "TKBPostback",
                    }],
                },
            },
        },
    }
    res = requests.post(req_mess, json = request_body)

import firebase_admin
from firebase_admin import firestore
from firebase_admin import credentials

def exec():
    global cfgData, req_gs, req_att, req_mess, req_noti
    cfgData = json.load(open('cfg.json', 'r', encoding='utf-8'))
    req_gs = f"https://script.google.com/macros/s/{cfgData['google_script_id']}/exec"
    req_att = f"https://graph.facebook.com/v13.0/me/message_attachments?access_token={cfgData['access_token']}"
    req_noti = f"https://script.google.com/macros/s/{cfgData['google_script_noti_id']}/exec"
    req_mess = f"https://graph.facebook.com/v14.0/me/messages?access_token={cfgData['access_token']}"
    firebase_admin.initialize_app(credentials.Certificate('key.json'))
    db = firestore.client()
    print(f"Đang lấy danh sách học sinh.")
    global max_val, val, bar, postNoti, postNotiMsg, usersList, debug
    if postNoti == "y":
        users = db.collection(u'RecurNoti').where(u'Enable', u'==', 1).stream()
        for user in users:
            usersList.append(user.to_dict()["RNToken"])
    max_val = (sheet.max_column - 2)*5 + len(usersList)
    val = -1
    bar = progressbar.ProgressBar(max_value=max_val,redirect_stdout=True, widgets=widgets)
    if debug == False:
        TotalArray = []
        with ThreadPoolExecutor(max_workers=10) as executor:
            for result in executor.map(execClass, range(2,sheet.max_column)):
                TotalArray.append(result)
                print(f"{len(TotalArray)}/{str(sheet.max_column-2)} | {result[0]} | Thêm thành công!")
        for i in range(2,sheet.max_column):
            os.remove(f"imgGen/index{i}.html")
        # === Phase 4: Cập nhật dữ liệu GG Sheets ================================ #
        print(f"Toàn bộ | Đang cập nhật dữ liệu GG Sheets")
        req = {
            "mode": 1,
            "data": TotalArray
        }
        retries = 0
        while retries<5:
            try:
                res = requests.post(req_gs, json.dumps(req))
                if res.status_code == 200:
                    break
            except:
                retries += 1
                print(f"Toàn bộ | Đang cập nhật dữ liệu GG Sheets (Lần {retries}/5)") 
        for i in range(2,sheet.max_column):
            updateProgBar("")
            time.sleep(2/(sheet.max_column-2))
    
    # === Phase 5: Thông tin cho Học sinh ================================ #
    if postNoti == "n":
        return
    cnt12 = 0
    with ThreadPoolExecutor(max_workers=int(cfgData['max_worker_count'])) as executor:
        for result in executor.map(execNoti, range(len(usersList))):
            cnt12 += 1
            print(f"{cnt12}/{len(usersList)} | Gửi thành công!")

def check():
    global today, wb, sheet, postNoti, postNotiMsg
    if os.path.exists('cfg.json') == False:
        print(bcolors.FAIL + "Không tìm thấy file cấu hình." + bcolors.ENDC)
        exit(0)
    
    if os.path.exists('key.json') == False:
        print(bcolors.FAIL + "Không tìm thấy file key." + bcolors.ENDC)
        exit(0)

    dataPath = input("Nhập đường dẫn file Excel .xlsx (kéo file vào): ")

    if os.path.exists(dataPath) == False or dataPath.endswith(".xlsx") == False:
        print(bcolors.FAIL + "File Excel không hợp lệ. File phải có dạng .xlsx." + bcolors.ENDC)
        exit(0)

    try:
        wb = openpyxl.load_workbook(dataPath)
        sheet = wb.worksheets[0]
    except:
        print(bcolors.FAIL + "Không thể mở file. Hãy kiểm tra lại file đầu vào." + bcolors.ENDC)
        exit(0)

    if sheet.cell(3,1).value != "Ngày":
        print(bcolors.FAIL + "File vào không hợp lệ. Đang thử thêm/bớt hàng vào đầu." + bcolors.ENDC)
        sheet.insert_rows(1,5)
        deleted = 0
        while sheet.cell(3,1).value != "Ngày" and deleted < 15:
            sheet.delete_rows(sheet.min_row, 1)
            deleted += 1
        if sheet.cell(3,1).value != "Ngày":
            print(bcolors.FAIL + "File vào không hợp lệ. Hãy đảm bảo hàng thứ ba trong sheet đầu tiên bao gồm tên lớp." + bcolors.ENDC)
            exit(0)

    print(bcolors.OKGREEN + "File hợp lệ. Kiểm tra thành công." + bcolors.ENDC)
    
    try:
        today = wb.worksheets[1].cell(2,6).value.date().strftime("%d/%m/%Y")
    except:
        today = wb.worksheets[1].cell(2,6).value
    print("  Số lớp:", str(sheet.max_column-2))
    print("Hiệu lực:", str(today))
    today2 = input("Nhập ngày mới nếu bị sai (Enter nếu đúng): ")
    if today2 != "":
        print("Hiệu lực:", str(today2))
        today = today2
    postNoti = input("Gửi thông báo đến học sinh? (y/n, mặc định là n): ")
    if postNoti == "":
        postNoti = "n"
    postNoti = postNoti.lower()
    if postNoti == "y":
        postNotiMsg = input("Ghi chú đến học sinh: ")
        print(f"Ghi chú đến HS: {postNotiMsg}")
        if postNotiMsg != "":
            postNotiMsg = "\nGhi chú: " + postNotiMsg
        else:
            postNotiMsg = "!"

check()
input("Nhấn Enter để cập nhật TKB.")
exec()
