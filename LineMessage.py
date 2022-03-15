import requests

try:
    TARGET_URL = 'https://notify-api.line.me/api/notify'
    TOKEN = 'HyzsVdhFD7USNTk4YsB21GXZvtPssAzPER9kTb0j7Xw'  # 발급받은 토큰
    headers = {'Authorization': 'Bearer ' + TOKEN}
    data = {'message': '파이썬에서 메시지를 보냅니다. 1 2 3!!'}

    requests.post(TARGET_URL, headers=headers, data=data)

except Exception as ex:
    print(ex)