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

# testImagePath = os.path.join(currentDir, "source", "test.png")
# testImagePath_output = os.path.join(currentDir, "source", "test111.png")
testPdfPath = os.path.join(currentDir, "source", "xinqianli.pdf")
# testPdfPath_output = os.path.join(currentDir, "source", "图纸111.pdf")
testPptPath = os.path.join(currentDir, "source", "PDF图文.pptx")
# testPptPath_output = os.path.join(currentDir, "source", "PDF图文111.pptx")
testXslxPath = os.path.join(currentDir, "source", "sheet.xlsx")

defaultConfigData = {
    "OCR_language": ["en","ru"],
    "OCR_minConfidence": 0.6,
    "OCR_fontSize": 0.7,
    "OCR_textPadding": True,
    "translate_to": "Chinese"
}


def extractTexts(filePath):
    if filePath.endswith(".pptx"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromPptx(filePath)
    elif filePath.endswith(".xlsx"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromXlsx(filePath)
    elif filePath.endswith(".pdf"):
        jsonMeta, jsonContent = TextsHandler.extractTextsFromPdf(filePath)
    return jsonMeta, jsonContent

def importTexts(jsonMeta):
    meta = TextsHandler._read_json(jsonMeta)
    filePath = meta["originalFilePath"]
    print(f"\nimporting texts from: {jsonMeta} to {filePath}")
    if filePath.endswith(".pptx"):
        TextsHandler.importTextsToPptx(jsonMeta)
    elif filePath.endswith(".xlsx"):
        TextsHandler.importTextsToXlsx(jsonMeta)
    elif filePath.endswith(".pdf"):
        TextsHandler.importTextsToPdf(jsonMeta)

def extractImages(filePath):
    if filePath.endswith(".pptx"):
        jsonMeta = ImageHandler.extractImagesFromPPTX(filePath)
    elif filePath.endswith(".xlsx"):
        jsonMeta = ImageHandler.extractImagesFromXLSX(filePath)
    elif filePath.endswith(".pdf"):
        jsonMeta = ImageHandler.extractImagesFromPDF(filePath)
    return jsonMeta

def importImages(jsonMeta):
    meta = TextsHandler._read_json(jsonMeta)
    filePath = meta["originalFilePath"]
    print(f"\nimporting images from: {jsonMeta} to {filePath}")
    if filePath.endswith(".pptx"):
        ImageHandler.importImagesToPPTX(jsonMeta)
    elif filePath.endswith(".xlsx"):
        ImageHandler.importImagesToXLSX(jsonMeta)
    elif filePath.endswith(".pdf"):
        ImageHandler.importImagesToPDF(jsonMeta)



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
    # ListCaller.translateJsonFile(jsonContent)
    writejsonToImage(jsonMeta)
    importImages(jsonPath)

def translateTextsFromFile(filePath):
    jsonMeta, jsonContent = extractTexts(filePath)
    ListCaller.translateJsonFile(jsonContent)
    importTexts(jsonMeta)


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
            ListCaller.targetLanguage = configData["translate_to"]
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

init()
userInput = input("\ngive me a file: ")
translateTextsFromFile(userInput)
# ImageTextWriter.writeJsonToImages("C:\\Users\\Administrator\\Desktop\\tools\\others\\MoonTranslator\\source\\chinese_extraction\\chinese_extraction_OCR.json")

