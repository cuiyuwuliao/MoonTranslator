import pyperclip
import time
import keyboard
import openai 
import sys
import os

curretDir = ""
if getattr(sys, 'frozen', False):  # If bundled by PyInstaller
    curretDir = os.path.dirname(sys.executable)
else:
    curretDir = os.path.dirname(os.path.abspath(__file__))

paidKey = "sk-R74koTKruRMkGXjogVbnPTCUzMkks2FeY76x3c0pJpBUuA0o"
client = openai.OpenAI(
    # defaults to os.environ.get("OPENAI_API_KEY")
    api_key="sk-WpCb6vXleVk6djvQICFygFTzv3B7GFvCEIN9YSV1C9ydNKW5",
    base_url="https://api.chatanywhere.tech/v1"
)

# Set this to True for standard mode, False for tray mode
STANDARD_MODE = True
FLUSH_MODE = True

copyStrike = {"count": 0, "lastStrikeTime":0}
lastRefreshTime = 0
def custom_print(*args, **kwargs):
    flush = FLUSH_MODE
    try:
        print(*args, flush=flush, **kwargs)
    except:
        print("!_内容包含不支持的编码, 无法显示!_")

def append_to_file(filename, text):
    file_path = os.path.join(curretDir, filename)
    with open(file_path, "a", encoding="utf-8") as f:
        f.write(text + "\n")

def send_prompt(prompt, model="gpt-3.5-turbo", max_tokens=3000):
    custom_print(f"!_翻译中<*_*>...!_")
    try:
        completion = client.chat.completions.create(
            model="gpt-3.5-turbo",
            max_tokens=max_tokens,
            messages=[
                {"role": "user","content":"You are a translator who provides precise and accurate translation, your response must starts with \" 中文意思为:\""},
                {
                    "role": "user",
                    "content": f"translate into Chinese: {prompt}",
                },
            ],
        )
        result = completion.choices[0].message.content
        custom_print(f"!_{result}!_")
        append_to_file("log_translations.txt", f"!_原文:{prompt}\n译文:{result}!_")
    except Exception as e:
        custom_print(f"!_错误: {e}!_")

def onCopy():
    global lastRefreshTime, copyStrike
    time.sleep(0.01)#等开Windows复制好内容
    current_text = pyperclip.paste()
    #连续复制四次启动翻译
    activate = False
    lastStrikeTime = copyStrike["lastStrikeTime"]
    currentTime = time.time()
    copyStrike["lastStrikeTime"] = currentTime
    if currentTime - lastStrikeTime < 0.5:
        copyStrike["count"] += 1
        if copyStrike["count"] > 2:
            activate = True
            copyStrike["count"] = 0
    else:
        copyStrike["count"] = 0

    if(currentTime - lastRefreshTime > 1800): #每半个小时强制刷新一次
        lastRefreshTime = currentTime
    if activate:
        send_prompt(current_text)
        
if STANDARD_MODE:
    keyboard.add_hotkey('ctrl+c', onCopy)
    keyboard.wait()
 