import json
import time

from selenium import webdriver

browers = webdriver.Chrome()
url = 'https://bj.meituan.com'
header = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/88.0.4324.190 Safari/537.36'
}
browers.get(url)
time.sleep(120)

list = browers.get_cookies()
print(list)
# with open('../data/cookies.json','w',encoding='utf-8') as file:
#     file.write(json.dumps(list))

with open('../data/cookies.json', 'w', encoding='utf-8') as file:
    json.dump(list, file, indent=2, ensure_ascii=False)

browers.close()