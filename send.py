import requests
import base64
import hashlib

# 图片base64码
with open("a.png", "rb") as f:
    base64_data = str(base64.b64encode(f.read()), 'utf8')
    f.close
# base64.b64decode(base64data)
print(base64_data)

# 图片的md5值
file = open("a.png", "rb")
md = hashlib.md5()
md.update(file.read())
res1 = md.hexdigest()
print(res1)

# 企业微信机器人发送消息
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
                "title": "不要忘记报名呀",
                "description": "点击报名考试",
                "url": "http://syjc.jtzyzg.org.cn/SYJC/LEAP/syjc/html/index.html",
                "picurl": "https://1.tel.jtzyzg.org.cn/allvideos/yxy/tiadlongrise/ltpu/images/syjc/title.png"
            }
        ]
    }
}
r = requests.post(url, headers=headers, json=data2)
print(r.text)
