print("启动中, 请稍等....")
print()
import TextsHandler
from OCRManager import ImageTextWriter
import ImageHandler
import ListCaller
import os
import sys
import json
import time

currentDir = ""
currentPath = ""
if getattr(sys, 'frozen', False):
    currentDir = os.path.dirname(sys.executable)
    currentPath = os.path.abspath(sys.executable)
else:
    currentDir = os.path.dirname(os.path.abspath(__file__))
    currentPath = os.path.abspath(__file__)
noTranslate = True
OCRlanguage = []

defaultConfigData = {
    "LLM_key": "sk-WpCb6vXleVk6djvQICFygFTzv3B7GFvCEIN9YSV1C9ydNKW5",
    "LLM_model": "gpt-4o-mini",
    "LLM_url": "https://api.chatanywhere.tech/v1", 
    "LLM_maxTries": 10, 
    "LLM_chunkSize": 1500,
    "OCR_minConfidence": 0.6,
    "OCR_fontSize": 0.7,
    "OCR_textPadding": True,
    "OCR_fontFile": "Chinese.ttf",
    "OCR_language": ["en","ru"],
    "translate_to": "Chinese"
}

def rename_file(original_file_path, new_file_name):
    if not os.path.isfile(original_file_path):
        return f"Error: The file '{original_file_path}' does not exist."
    directory = os.path.dirname(original_file_path)
    file_extension = os.path.splitext(original_file_path)[1]
    new_file_path = os.path.join(directory, f"{new_file_name}{file_extension}")
    try:
        os.rename(original_file_path, new_file_path)
        return f"File renamed successfully to: {new_file_path}"
    except Exception as e:
        return f"Error: {str(e)}"
def get_file_name_without_extension(file_path):
    base_name = os.path.basename(file_path)
    file_name_without_extension = os.path.splitext(base_name)[0]
    return file_name_without_extension

def extractTexts(filePath):
    if filePath.endswith(".pptx"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromPptx(filePath)
    elif filePath.lower().endswith(".xlsx"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromXlsx(filePath)
    elif filePath.lower().endswith(".pdf"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromPdf(filePath)
    elif os.path.isdir(filePath) or filePath.lower().endswith(('.jpeg', '.jpg', '.png')):
        jsonMeta,jsonContent = runOCR(filePath)
    return jsonMeta, jsonContent

def importTexts(jsonMeta):
    meta = TextsHandler._read_json(jsonMeta)
    filePath = meta["originalFilePath"]
    print(f"\nimporting texts from: {jsonMeta} to {filePath}")
    if filePath.lower().endswith(".pptx"):
        outputFileName = TextsHandler.importTextsToPptx(jsonMeta)
    elif filePath.lower().endswith(".xlsx"):
        outputFileName = TextsHandler.importTextsToXlsx(jsonMeta)
    elif filePath.lower().endswith(".pdf"):
        outputFileName = TextsHandler.importTextsToPdf(jsonMeta)
    return outputFileName

def extractImages(filePath, isRunOCR = False):
    if filePath.lower().endswith(".pptx"):
        jsonMeta = ImageHandler.extractImagesFromPPTX(filePath)
    elif filePath.lower().endswith(".xlsx"):
        jsonMeta = ImageHandler.extractImagesFromXLSX(filePath)
    elif filePath.lower().endswith(".pdf"):
        jsonMeta = ImageHandler.extractImagesFromPDF(filePath)
    if isRunOCR:
        runOCR(os.path.dirname(jsonMeta))
    return jsonMeta

def importImages(jsonMeta):
    meta = TextsHandler._read_json(jsonMeta)
    filePath = meta["originalFilePath"]
    print(f"\nimporting images from: {jsonMeta} to {filePath}")
    if filePath.endswith(".pptx"):
        outputFileName = ImageHandler.importImagesToPPTX(jsonMeta)
    elif filePath.endswith(".xlsx"):
        outputFileName = ImageHandler.importImagesToXLSX(jsonMeta)
    elif filePath.endswith(".pdf"):
        outputFileName = ImageHandler.importImagesToPDF(jsonMeta)
    return outputFileName



def runOCR(filePath):
    ImageTextWriter.loadIamgeSource(filePath)
    jsonMeta, jsonContent = ImageTextWriter.exportJson()
    return jsonMeta, jsonContent

def writejsonToImage(jsonResult):
    ImageTextWriter.writeJsonToImages(jsonResult)

def translateImagesFromFile(filePath):
    jsonPath = extractImages(filePath)
    imageDir = os.path.split(jsonPath)[0]
    jsonMeta, jsonContent = runOCR(imageDir)
    if not noTranslate:
        print("###Begin Images Translation###")
        ListCaller.translateJsonFile(jsonContent)
    writejsonToImage(jsonMeta)
    outputFileName = importImages(jsonPath)
    return outputFileName

def translateTextsFromFile(filePath):
    jsonMeta, jsonContent = extractTexts(filePath)
    if not noTranslate:
        print("###Begin Texts Translation###")
        ListCaller.translateJsonFile(jsonContent)
    outputFileName = importTexts(jsonMeta)
    return outputFileName

def translateImage(filePath):
    jsonMeta,jsonContent = runOCR(filePath)
    if not noTranslate:
        print("###Begin Images Translation###")
        ListCaller.translateJsonFile(jsonContent)
    writejsonToImage(jsonMeta)
    return jsonMeta

def translateFile(filePath):
    if os.path.isdir(filePath) or filePath.lower().endswith(('.jpeg', '.jpg', '.png')):
        outputFileName = translateImage(filePath)
    else:
        textTranslated = translateTextsFromFile(filePath)
        allTranslated = translateImagesFromFile(textTranslated)
        os.remove(textTranslated)
        rename_file(allTranslated, f"{get_file_name_without_extension(filePath)}_translation")
    return allTranslated

def init():
    global defaultConfigData, OCRlanguage
    configFile = os.path.join(currentDir, "config.json")
    try:
        with open(configFile, 'r', encoding='utf-8') as file:
            configData = json.load(file)
            OCRlanguage = configData["OCR_language"]
            ImageTextWriter.setLanguage(configData["OCR_language"])
            ImageTextWriter.minConfidence = configData["OCR_minConfidence"]
            ImageTextWriter.FontSizeRatio = configData["OCR_fontSize"]
            ImageTextWriter.textPadding = configData["OCR_textPadding"]
            ImageTextWriter.fontFile = configData["OCR_fontFile"]
            ListCaller.targetLanguage = configData["translate_to"]
            ListCaller.setLLM(url=configData["LLM_url"], key=configData["LLM_key"])
            ListCaller.maxTries = configData["LLM_maxTries"]
            ListCaller.model = configData["LLM_model"]
            ListCaller.chunkSize = configData["LLM_chunkSize"]
    except Exception as e:
        if os.path.exists(configFile):
            os.remove(configFile)
        print(f"\n! config.json is corrupted or does not exists, creating a defualt config.json....")
        with open(configFile, 'w', encoding='utf-8') as file:
            json.dump(defaultConfigData, file, ensure_ascii=False, indent=2)
        print("\n请确保config.json中的配置正确之后再重新运行")
        os.startfile(configFile)
        time.sleep(5)
        sys.exit()


def createWriteJsonBat(folderPath):
    if not os.path.isdir(folderPath):
        folderPath = os.path.dirname(folderPath)
    bat_content = f"""@echo off
    {currentPath} "{"writeJson"}" "{folderPath}"
    pause
    """
    bat_file_name = os.path.join(folderPath, "recoverFile.bat")
    if os.path.exists(bat_file_name):
        return bat_file_name
    with open(bat_file_name, 'w') as bat_file:
        bat_file.write(bat_content)
    return bat_file_name


if __name__ == "__main__":
    init()
    arg = None
    filePath = None
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if len(sys.argv) > 2:
            filePath = sys.argv[2]
    if arg == None:
        arg = input("argument: ")


    if arg == "extract":
        while filePath == None or not os.path.exists(filePath):
            filePath = input("拖入要提取的文件(拖入文件夹时, 会提取文件夹内所有图片的文字):\n")
        jsonMeta, jsonContent = extractTexts(filePath)
        createWriteJsonBat(jsonMeta)
        if not os.path.isdir(filePath) or filePath.lower().endswith(('.jpeg', '.jpg', '.png')):
            extractImages(filePath, isRunOCR=True)
            

    if arg == "translate":
        while filePath == None or not os.path.exists(filePath):
            filePath = input("拖入要翻译的文件(拖入文件夹时, 会翻译文件夹下所有的图片):\n")
        translateFile(filePath)

    if arg == "OCR":
        while filePath == None or not os.path.exists(filePath):
            filePath = input(f"拖入要提取文字的图片或文件夹(当前可以识别:{",".join(OCRlanguage)}):\n")
        jsonMeta, jsonContent = runOCR(filePath)
        createWriteJsonBat(jsonMeta)

    if arg == "writeJson":
        while filePath == None or not os.path.exists(filePath):
            filePath = input("拖入meta文件(或文件夹)将提取的内容写回原文档:\n")
        pathList = []
        if os.path.isdir(filePath):
            for filename in os.listdir(filePath):
                full_path = os.path.join(filePath, filename)
                if os.path.isfile(full_path):
                    pathList.append(full_path)
        else:
            pathList.append(filePath)
        uniqueList = []
        for path in pathList:  
            if path.endswith("_OCR_content.json") or path.endswith("_OCR.json"):
                path = path.replace("_OCR_content.json", "_OCR.json")
                if path not in uniqueList:
                    writejsonToImage(path)
                    uniqueList.append(path)
            elif path.endswith("_texts_content.json") or path.endswith("_texts_meta.json"):
                path = path.replace("_texts_content.json", "_texts_meta.json")
                if path not in uniqueList:
                    importTexts(path)
                    uniqueList.append(path)
            elif path.endswith("_images_meta.json"):
                if path not in uniqueList:
                    importImages(path)
                    uniqueList.append(path)