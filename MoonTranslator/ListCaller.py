
import openai 
import json

paidKey = "sk-R74koTKruRMkGXjogVbnPTCUzMkks2FeY76x3c0pJpBUuA0o"
freeKey = "sk-WpCb6vXleVk6djvQICFygFTzv3B7GFvCEIN9YSV1C9ydNKW5"
client = openai.OpenAI(
    # defaults to os.environ.get("OPENAI_API_KEY")
    api_key=paidKey,
    base_url="https://api.chatanywhere.tech/v1"
)
targetLanguage = "Chinese"

def split_list_by_length(input_list, max_chars=1000):
    if not input_list:
        return []
    chunks = []
    current_chunk = []
    current_length = 2  # Start with 2 for the opening and closing brackets of JSON array
    for item in input_list:
        # Convert item to JSON string
        item_str = json.dumps(item)
        item_length = len(item_str)
        # Account for comma if not the first item in chunk
        separator_length = 1 if current_chunk else 0
        # Check if adding this item would exceed the limit
        if current_chunk and (current_length + separator_length + item_length) > max_chars:
            chunks.append(current_chunk)
            current_chunk = [item]
            current_length = 2 + item_length  # brackets + item
        else:
            current_chunk.append(item)
            current_length += separator_length + item_length
    
    if current_chunk:
        chunks.append(current_chunk)
    return chunks

def send_prompt(list, model="gpt-3.5-turbo", max_tokens=3000):
    global targetLanguage
    try:
        completion = client.chat.completions.create(
            model=model,
            max_tokens=max_tokens,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system",
                 "content":
f"You are a professional translator specializing in JSON data translation. Your task is to translate all string elements from the input JSON list into {targetLanguage} while preserving the exact same structure and length\n"
+"""
CRITICAL RULES:
1. Return ONLY a JSON object with this exact format: {"translation": ["translated_text1", "translated_text2", ...]}
2. The output list must have exactly the same number of elements as the input list
3. Do not translate numbers and special symbols
4. The elements are not independent from each other, Try Maintaining the context and natural phrasing
5. Do not add escape sequences

Example input: ["我喜欢你", "的勇气和", "思想", ""]
Example output: {"translation": ["I like your", "bravery and", "your mind", ""]}"""},
                {
                    "role": "user",
                    "content": f"{json.dumps(list)}"
                }
            ]
        )
        result = completion.choices[0].message.content

        try:
            result_ls = json.loads(result)
            result_ls = result_ls["translation"]
            print("-------------------------------------")
            print(f"Part Complete:\noriginal list count: {len(list)}\nresult list count: {len(result_ls)}")
            print({"original":list})
            print(result)
            print("-------------------------------------")
            if len(list) != len(result_ls):
                print("! invalid translation\nretrying...")
                result_ls = send_prompt(list)
        except Exception as e:
                print(f"! exception:{e}\nretrying...")
                result_ls = send_prompt(list)
        return result_ls
    except Exception as e:
        print(f"!_错误: {e}!_")


def translateList(list):
    chunks = split_list_by_length(list)
    chunks_result = []
    for chunk in chunks:
        chunks_result+=send_prompt(chunk)
    print(f"\nTranslation Complete:\noriginal list count: {len(list)}\nresult list count: {len(chunks_result)}")
    return chunks_result

def translateJsonFile(jsonPath):
    with open(jsonPath, 'r', encoding='utf-8') as file:
        list = json.load(file)
    promptList = []
    for item in list:
        promptList.append(item["text"])
    chunks = split_list_by_length(promptList)
    chunks_result = []
    index = 1
    for chunk in chunks:
        print(f"\n$ Translating: {index}/{len(chunks)}")
        result = send_prompt(chunk)
        chunks_result += result
        index += 1
    print(f"\nTranslation Complete")
    index = 0
    for item in list:
        list[index]["text"] = chunks_result[index]
        index += 1
    with open(jsonPath, 'w', encoding='utf-8') as json_file:
        json.dump(list, json_file, ensure_ascii=False, indent=2)
    return list
    