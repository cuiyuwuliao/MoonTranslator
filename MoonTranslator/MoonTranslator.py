import TextsHandler
from OCRManager import ImageTextWriter
import ImageHandler
import os
import sys
import ListCaller
import json
import time

currentDir = ""
if getattr(sys, 'frozen', False):
    currentDir = os.path.dirname(sys.executable)
else:
    currentDir = os.path.dirname(os.path.abspath(__file__))
noTranslate = True

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


def extractTexts(filePath):
    if filePath.endswith(".pptx"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromPptx(filePath)
    elif filePath.lower().endswith(".xlsx"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromXlsx(filePath)
    elif filePath.lower().endswith(".pdf"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromPdf(filePath)
    elif filePath.lower().endswith(('.jpeg', '.jpg', '.png')):
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

def extractImages(filePath):
    if filePath.lower().endswith(".pptx"):
        jsonMeta = ImageHandler.extractImagesFromPPTX(filePath)
    elif filePath.lower().endswith(".xlsx"):
        jsonMeta = ImageHandler.extractImagesFromXLSX(filePath)
    elif filePath.lower().endswith(".pdf"):
        jsonMeta = ImageHandler.extractImagesFromPDF(filePath)
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
        outputFileName = translateTextsFromFile(filePath)
        outputFileName = translateImagesFromFile(outputFileName)
    return outputFileName

def init():
    global defaultConfigData
    configFile = os.path.join(currentDir, "config.json")
    try:
        with open(configFile, 'r', encoding='utf-8') as file:
            configData = json.load(file)
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



# init()
# userInput = input("\ngive me a file: ")
# translateImagesFromFile(userInput)

# init()
# userInput = input("\ngive me a file: ")
# translateTextsFromFile(userInput)
# ImageTextWriter.writeJsonToImages("C:\\Users\\Administrator\\Desktop\\tools\\others\\MoonTranslator\\source\\chinese_extraction\\chinese_extraction_OCR.json")

init()
userInput = input("\ngive me a file: ")
translateFile(userInput)