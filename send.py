import requests
import base64
import hashlib

# 图片base64码
with open("OIP-C.jpg", "rb") as f:
    base64_data = str(base64.b64encode(f.read()), 'utf8')
    f.close
# base64.b64decode(base64data)
print(base64_data)

# 图片的md5值
file = open("OIP-C.jpg", "rb")
md = hashlib.md5()
md.update(file.read())
res1 = md.hexdigest()
print(res1)

# 企业微信机器人发送消息
url = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=0e607fb7-72f2-421f-91a6-1425c78ad28c'
headers = {"Content-Type": "application/json"}
data1 = {
    "msgtype": "image",
    "image": {
        "base64": base64_data,
        "md5": res1
    }
}

data2 = {
    "msgtype": "news",
    "news": {
        "articles": [
            {
                "title": "这个年纪你睡得着觉?!",
                "description": "快工作!",
                "url": r"http://sdgsnavi.sensorcmd.com/#/entrance",
                "picurl": "https://tse4-mm.cn.bing.net/th/id/OIP-C.yVSnGS_5v3Al426NDGPKxQAAAA?w=171&h=180&c=7&r=0&o=5&pid=1.7"
            }
        ]
    }
}
r = requests.post(url, headers=headers, json=data2)
print(r.text)
