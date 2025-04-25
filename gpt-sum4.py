import pandas as pd
import openai
from openpyxl import Workbook
import re
import requests
import json
import time



# 设置您的OpenAI API密钥
openai.api_key = "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"


# 定义读取Excel文件函数
def read_excel(file_name):
    data = pd.read_excel(file_name, header=None, engine='openpyxl')
    return data

# 定义写入Excel文件函数
def write_excel(file_name, data, headers):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(headers)
    for row in data:
        processed_row = [str(item) if isinstance(item, list) else item for item in row]  # 处理列表并转换为字符串
        worksheet.append(processed_row)
    workbook.save(file_name)



# 定义调用OpenAI函数
def call_openai_api(input_text):
     jscode = '''
         {
            "features": [
                {
                    "Dimension": "Dimension1",
                    "Reason":"Reason1",
                    "Importance_Score": Importance_Score1,
                    "Sentiment_Score":Sentiment_Score1,
                    "Phrases": [
                        "phrases_a1",
                        "phrases_a2"
                    ]
                },
                {
                    "Dimension": "Dimension2",
                    "Reason":"Reason2",
                    "Importance_Score": Importance_Score2,
                    "Sentiment_Score":Sentiment_Score2,
                    "Phrases": [
                        "phrases_b1",
                        "phrases_b2"
                    ]
                }
            ]
        }
        '''
     try:
        openai.api_key = "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"

        response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        max_tokens=2000,
        temperature=0.7,
        messages=[
                {"role": "system", "content": "You are an experienced product operator, and you can identify the features of a product from external comments."},
                {"role": "user", "content":f"""
                        Your task is to perform the following steps and ultimately provide a report on the points that users are most concerned about to the product department:

                        1- Read the text inside the <a></a> tags. These texts are user comments or media articles related to egg yolk pastry. Please thoroughly understand.
                        2- Summarize the dimensions of 'Egg-Yolk Puff' that users care about the most from the text and output the results according to the following rules:
                            ---
                            1. The dimensions should be expressed in Chinese.
                            2. Keep the dimensions concise, not exceeding 4 characters.
                            ---
                        3- Output the reasons for summarizing this dimension in Chinese.
                        4- Assign a score to each dimension based on its importance to users, ranging from 0 to 10, with higher scores indicating higher importance.
                        5- Assess the positive or negative sentiment of the dimension based on the semantic meaning of the original text, and assign a score ranging from 0 to 10. A higher score indicates a more positive sentiment.
                        6- Extract relevant phrases or word combinations related to each dimension from the original text, following these rules:
                            ---
                            1. Extract Chinese content.
                            2. Give priority to extracting phrases with verb-object or attributive-modifier structures.
                            3. Limit the extracted phrases to a maximum of 5 characters.
                            4. The extracted phrases should reflect  feelings or experiences of users.
                            5. The extracted phrases should express positive and favorable sentiments.
                            ---
                        7- Finally, output the results in JSON format, as shown in the example below:
                            {jscode}
                       
                        Text to read: <a>{input_text}</a>

                        """},
            ]
        )

        result = response["choices"][0]["message"]["content"]
        # print(result)
        return result
     
     except Exception as exc: #捕获异常后打印出来
        print(exc)

# 调用openAI输出json  
def json_output(data):
    # 将提取文本传给openai
    raw_string = call_openai_api(data)
    print('raw:')
    print(raw_string)


    if raw_string:
        # 提取 JSON 部分并去除非法字符
        try:
            json_string = re.search(r'{[\s\S]*}', raw_string).group()
        except AttributeError:
            json_string = "{'features': [{'Dimension': '0','Reason':'0', 'Importance_Score': 0,'Sentiment_Score':0, 'Phrases': '0'}]}"
        cleaned_string = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', json_string)
        print('clean')
        print(cleaned_string)

        # 解析 JSON 数据
        try:
            arr = json.loads(cleaned_string)
            print('arr:')
            print(arr)
        except json.JSONDecodeError as e:
            print("JSON 解析出错:", e)
            arr = {'features': [{'Dimension': '0','Reason':'0', 'Importance_Score': 0,'Sentiment_Score':0, 'Phrases': '0'}]}
            print(arr)


        # 提取features数组
        features = arr["features"]
        print('feature:')
        print(features)
        return features
    else:
        return [{'Dimension': '0','Reason':'0', 'Importance_Score': 0,'Sentiment_Score':0, 'Phrases': '0'}]
    

        

# 定义主函数
def main():
    # 读取源Excel文件
    input_file = "C:/Users/chenchuxiao/Desktop/gpt-md/input_file_xhs_dhs_2.xlsx"
    data = read_excel(input_file)

    # 创建一个空列表用于存储数据
    new_data = [] 

    #初始化excel表头
    headers = ["num","Dimension", "Reason","Importance_Score","Sentiment_Score", "Phrases"]

    # 遍历每一行数据，并调用接口获取 JSON 数据并存入 new_data 列表
    for index, row in data.iterrows():
        original_data = ' '.join(['{}'.format(cell).strip() for cell in row if pd.notnull(cell)])
        # original_data = row.tolist()  # 当前行的原始数据


        # 调用接口并获取 JSON数组 数据
        json_data = json_output(original_data)  # 调用你的接口函数，并传递原始数据

       

        # 如果 JSON 数据不为空，则将其转换为二维列表并添加到 new_data 列表中
        if json_data:
            for feature in json_data:
                dimension = feature["Dimension"]
                reason = feature["Reason"]
                importance_score = feature["Importance_Score"]
                sentiment_score = feature["Sentiment_Score"]
                phrases = feature["Phrases"]
                new_data.append([index+1,dimension,reason,importance_score,sentiment_score,phrases])

        

        # 打印日志
        print(f"Processed row {index + 1}")
        # print(json_data)
        time.sleep(21)

    # 如果 new_data 列表不为空，则创建新的 Excel 文件并写入数据
    
    output_file = "C:/Users/chenchuxiao/Desktop/gpt-md/output_file_xhs_dhs_3.xlsx"
    if new_data:
        # 在new_data中添加num列
        new_data_with_num = [row[:6] for row in new_data]
        write_excel(output_file, new_data_with_num, headers)
    else:
        print("No data to write.")


# 执行主函数
if __name__ == "__main__":
    main()
