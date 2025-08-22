import pyperclip
import time
import keyboard
import openai 
import sys
import os
import json
# Set this to True for standard mode, False for tray mode
STANDARD_MODE = True
FLUSH_MODE = True

def custom_print(*args, **kwargs):
    flush = FLUSH_MODE
    try:
        print(*args, flush=flush, **kwargs)
    except:
        print("!_内容包含不支持的编码, 无法显示!_")
custom_print("!_复制翻译器已启动, 复制一个文本四次即可翻译成中文!_")

curretDir = ""
if getattr(sys, 'frozen', False):  # If bundled by PyInstaller
    curretDir = os.path.dirname(sys.executable)
else:
    curretDir = os.path.dirname(os.path.abspath(__file__))

config_key = None
config_url = None
config_model = None

configFile = os.path.join(curretDir, "config.json")
try:
    with open(configFile, 'r', encoding='utf-8') as file:
        configData = json.load(file)
        OCRlanguage = configData["OCR_language"]
        config_url= configData["LLM_url"]
        config_key=configData["LLM_key"]
        config_model = configData["LLM_model"]
except Exception as e:
    custom_print("!_复制翻译器启动失败,请检查LLM配置_")
    os.startfile(configFile)
    time.sleep(5)
    sys.exit()

client = openai.OpenAI(
    # defaults to os.environ.get("OPENAI_API_KEY")
    api_key=config_key,
    base_url=config_url
)



copyStrike = {"count": 0, "lastStrikeTime":0}
lastRefreshTime = 0

def append_to_file(filename, text):
    file_path = os.path.join(curretDir, filename)
    with open(file_path, "a", encoding="utf-8") as f:
        f.write(text + "\n")

def send_prompt(prompt, model="gpt-4o-mini", max_tokens=3000):
    
    if prompt == "" or "!_" in prompt:
        custom_print(f"!_复制内容无效: {prompt}\n请重试, 请不要同时复制连续的感叹号和下划线!_")
    else:
        custom_print(f"!_翻译中<*_*>...!_")
    try:
        completion = client.chat.completions.create(
            model=model,
            max_tokens=max_tokens,
            messages=[
                {"role": "system","content":"You are a translator who provides precise and accurate translation, your response must starts with \" 译文:\""},
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
    time.sleep(0.1)#等开Windows复制好内容
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
 