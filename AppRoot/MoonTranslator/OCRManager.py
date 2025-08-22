from PIL import Image, ImageDraw, ImageFont
import easyocr
import os
import sys
import io
import json
import numpy as np
import concurrent.futures

currentDir = ""
if getattr(sys, 'frozen', False):
    currentDir = os.path.dirname(sys.executable)
else:
    currentDir = os.path.dirname(os.path.abspath(__file__))


def add_suffix_to_filename(file_path, suffix):
    # Split the file path into directory, base name, and extension
    directory, filename = os.path.split(file_path)
    name, ext = os.path.splitext(filename)
    
    # Create the new filename with the suffix
    new_filename = f"{name}{suffix}{ext}"
    
    # Join the new filename with the directory
    new_file_path = os.path.join(directory, new_filename)
    
    return new_file_path

def convert_numpy_types(data):
    if isinstance(data, dict):
        return {key: convert_numpy_types(value) for key, value in data.items()}
    elif isinstance(data, list):
        return [convert_numpy_types(item) for item in data]
    elif isinstance(data, np.integer):
        return int(data)  # Convert NumPy integer to Python int
    elif isinstance(data, np.floating):
        return float(data)  # Convert NumPy float to Python float
    else:
        return data  # Return the data as is if it's not a NumPy type

class ImageTextWriter:
    reader = None
    itemList = []
    defaultJsonPath = os.path.join(currentDir,"Ocr_Content.json")
    sessionName = None
    sessionDir = None
    minConfidence = 0.7 #最低的识别度, 不超过0.9的文字忽略
    FontSizeRatio = 0.7 #字体大小, 1代表和原本一样大
    textPadding = True
    fontFile = "Chinese.ttf"

    def __init__(self, inputPath = None, setOcr=True):
        self.inputPath = inputPath
        self.outputPath = inputPath
        self.baseName = None
        self.ImageObj = None
        self.Draw = None
        self.texts = []
        if isinstance(inputPath, str) and os.path.isfile(inputPath):
            self.ImageObj = Image.open(inputPath)
            self.Draw = ImageDraw.Draw(self.ImageObj)
            self.baseName = os.path.splitext(os.path.split(inputPath)[1])[0]
            if setOcr:
                self.setOcrResult()
        elif isinstance(inputPath, Image.Image):
            self.ImageObj = inputPath
            self.Draw = ImageDraw.Draw(self.ImageObj)
            if setOcr:
                self.setOcrResult()
        ImageTextWriter.itemList.append(self)

    def setLanguage(ls = ['en']):
        ImageTextWriter.reader = easyocr.Reader(ls) 

    def setOcrResult(self):
        ocrResult = ImageTextWriter.reader.readtext(self.inputPath)
        textIndex = 0
        for entry in ocrResult:
            BoundingBox = entry[0]
            confidence = entry[2] if len(entry) > 2 else 0
            text = entry[1]
            if confidence < ImageTextWriter.minConfidence:
                continue
            self.texts.append({"text":text, "confidence": confidence, "boundingBox": BoundingBox, "imagePath": self.inputPath if isinstance(self.inputPath, str) else None, "index":f"{self.baseName}_{textIndex}"})
            textIndex += 1

    def writeTextEntryToImage(ImageObj, Text, BoundingBox, FontSizeRatio = None):
        if not isinstance(FontSizeRatio, float):
            FontSizeRatio = ImageTextWriter.FontSizeRatio
        # Calculate the bounding box coordinates
        TopLeft = BoundingBox[0]
        BottomRight = BoundingBox[2]
        # Calculate width and height of the bounding box
        BoxWidth = abs(BottomRight[0] - TopLeft[0])
        BoxHeight = abs(BottomRight[1] - TopLeft[1])
        if BoxHeight < 20:
            BoxHeight = 20
        FontSize = int(BoxHeight * FontSizeRatio)
        Font = ImageFont.truetype(os.path.join(currentDir,ImageTextWriter.fontFile), FontSize)
        # Calculate text size
        if ImageObj.mode != 'RGB':
            ImageObj = ImageObj.convert('RGB')
        Draw = ImageDraw.Draw(ImageObj)
        TextBoundingBox = Draw.textbbox((TopLeft[0], TopLeft[1]), Text, font=Font)
        TextWidth = abs(TextBoundingBox[2] - TextBoundingBox[0])
        TextHeight = abs(TextBoundingBox[3] - TextBoundingBox[1])
        # Calculate position (centered in the bounding box)
        X = TopLeft[0]
        Y = TopLeft[1] + BoxHeight
        # Define text color
        TextColor = (255, 0, 255) 
        # Add text to image
        newFontSize = FontSize * (BoxWidth / TextWidth)
        Font = ImageFont.truetype(os.path.join(currentDir,ImageTextWriter.fontFile), newFontSize)
        if ImageTextWriter.textPadding:
            TextBoundingBox = Draw.textbbox((TopLeft[0], TopLeft[1]), Text, font=Font)
            TextWidth = TextBoundingBox[2] - TextBoundingBox[0]
            TextHeight = TextBoundingBox[3] - TextBoundingBox[1]
            Draw.rectangle([X, Y, X+TextWidth, Y+TextHeight], fill=(0,255,0))
        Draw.text((X, Y), Text, fill=TextColor, font=Font)

    def exportJson(jsonPath = None):
        if jsonPath == None:
            jsonPath = ImageTextWriter.defaultJsonPath
        ocrAttributes = []
        ocrTexts = []
        for item in ImageTextWriter.itemList:
            if not isinstance(item.baseName, str):
                return
            for itemText in item.texts:
                jsonData = convert_numpy_types(itemText)
                textId = itemText["index"]
                jsonData["text"] = textId
                ocrTexts.append({"Id":textId, "text":itemText["text"]})
                ocrAttributes.append(jsonData)
        with open(jsonPath, 'w', encoding='utf-8') as json_file:
            json.dump(ocrAttributes, json_file, ensure_ascii=False, indent=2)
        textsFilePath = add_suffix_to_filename(jsonPath, "_content")
        with open(textsFilePath, 'w', encoding='utf-8') as json_file:
            json.dump(ocrTexts, json_file, ensure_ascii=False, indent=2)
        return jsonPath, textsFilePath

    def getBytesObject(self):
        imgByteArray = io.BytesIO()
        self.ImageObj.save(imgByteArray, format="PNG")
        imgByteArray.seek(0)  
        return imgByteArray

    def loadIamgeSource(path):
        ImageTextWriter.itemList.clear()
        sessionName = None
        sessionDir = None
        def process_image(file_path):
            try:
                return ImageTextWriter(file_path)
            except Exception as e:
                print(f"Error processing {file_path}: {e}")
                return None
        if os.path.isfile(path):
            sessionName = os.path.splitext(os.path.split(path)[1])[0]
            sessionDir = os.path.split(path)[0]
            if path.lower().endswith(('.jpeg', '.jpg', '.png')):
                process_image(path)
        elif os.path.isdir(path):
            sessionName = os.path.basename(os.path.normpath(path))
            sessionDir = path
            # Collect all image files
            image_files = []
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.lower().endswith(('.jpeg', '.jpg', '.png')):
                        image_files.append(os.path.join(root, file))
            # Process images in parallel
            max_workers = min(8, os.cpu_count() * 2)  # Adjust based on your system
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Submit all tasks
                future_to_path = {executor.submit(process_image, img_path): img_path for img_path in image_files}
                # Process results as they complete
                for future in concurrent.futures.as_completed(future_to_path):
                    img_path = future_to_path[future]
                    print("---------------------------------")
                    try:
                        result = future.result()
                        if result:
                            print(f"OCR process complete: {img_path}")
                            print(result.texts)
                    except Exception as e:
                        print(f"Error processing {img_path}: {e}")
                print("---------------------------------")
        else:
            print("The provided path is neither a file nor a directory.")
            return
        ImageTextWriter.sessionName = sessionName
        ImageTextWriter.sessionDir = sessionDir
        ImageTextWriter.defaultJsonPath = os.path.join(sessionDir, f"{sessionName}_OCR.json")
        return ImageTextWriter.defaultJsonPath, 


    def writeJsonToImages(jsonPath = None):
        if jsonPath == None:
            jsonPath = ImageTextWriter.defaultJsonPath
        with open(jsonPath, 'r', encoding='utf-8') as file:
            result = json.load(file)
        with open(add_suffix_to_filename(jsonPath,"_content"), 'r', encoding='utf-8') as file:
            result_texts = json.load(file)
        imageObjects = []
        for data in result:
            textId = data["text"]
            for resultText in result_texts:
                if resultText["Id"] == textId:
                    text = resultText["text"]
                    break
            imagePath = data["imagePath"]
            if not any(imagePath in d.values() for d in imageObjects):
                imageObjects.append({"imagePath":imagePath, "imageObject":Image.open(imagePath)})
            boundingBox = data["boundingBox"]
            for imageObj in imageObjects:
                if imageObj["imagePath"] == imagePath:
                    ImageTextWriter.writeTextEntryToImage(imageObj["imageObject"], text, boundingBox)
                    print(f"$ writing: {text} -> {imagePath}")
                    break
        for imageObj in imageObjects:
            imageObj["imageObject"].save(imageObj["imagePath"])
            
