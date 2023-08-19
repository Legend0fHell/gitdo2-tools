import json
import os
import requests

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

def exec(imgPath):
    cfgData = json.load(open('cfg.json', 'r', encoding='utf-8'))
    req_att = f"https://graph.facebook.com/v13.0/me/message_attachments?access_token={cfgData['access_token']}"
    # === Phase 3: Đăng ảnh lên API Messenger ================================ #
    payload = {
        'message': (None, {"attachment": {"type": "image", "payload": {"is_reusable": 'true'}}}),
    }
    files = {
        'filedata': ('test', open(imgPath, 'rb'), 'image/png')
    }
    retries = 0
    while retries < 5:
        try:
            res = requests.post(req_att, data=payload, files=files)
            attID = json.loads(res.content)
            if res.status_code == 200:
                break
        except:
            retries += 1
    print("Thêm thành công! ID Messenger: ", attID["attachment_id"])


def pre():
    if os.path.exists('cfg.json') == False:
        print(bcolors.FAIL + "Không tìm thấy file cấu hình." + bcolors.ENDC)
        exit(0)
    imgPath = input("Nhập đường dẫn ảnh (kéo ảnh vào): ")
    input("Nhấn Enter để post ảnh.")
    exec(imgPath)
    print(f"\n{bcolors.OKGREEN}Hoàn thành.")

pre()
