import os
import base64
from openai import OpenAI
import re
import time
import logging
import fitz
import pandas as pd
import anthropic
import statistics
import docx
from PIL import Image
import pytesseract
import requests
import random
from google import genai
from google.genai import types
def calculate_mean(data):
    filtered_data = [x for x in data if isinstance(x, (int, float))]
    if len(filtered_data) > 2:
        filtered_data.remove(max(filtered_data))
        filtered_data.remove(min(filtered_data))
        return statistics.mean(filtered_data)
    else:
        return None 

def calculate_std_dev(data):
    filtered_data = [x for x in data if isinstance(x, (int, float))]
    if len(filtered_data) > 2:
        filtered_data.remove(max(filtered_data))
        filtered_data.remove(min(filtered_data))
        if len(filtered_data) < 2:
            return None
        return statistics.stdev(filtered_data)
    else:
        return None 

def setup_logging(model_name):
    log_filename = 'result.log'
    logging.basicConfig(
        filename=log_filename,
        level=logging.INFO,
        format='%(asctime)s %(levelname)s:%(message)s'
    )

def extract_answer(output):
    lines = output.strip().split('\n')
    if lines:
        last_line = lines[-1]
        fraction_match = re.search(r"(\d+)/(\d+)", last_line)
        if fraction_match:
            return int(fraction_match.group(1))
    numbers = re.findall(r"\d+", output)
    if numbers:
        return int(numbers[-1])
    return output

def read_and_encode_file(filename):
    with open(filename, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode("utf-8")

def list_files_in_directory(directory):
    return [os.path.join(directory, f) for f in os.listdir(directory) if os.path.isfile(os.path.join(directory, f))]

def extract_text_from_pdf(filename):
    doc = fitz.open(filename)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

def extract_text_from_docx(filename):
    doc = docx.Document(filename)
    text = ''
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_text_from_image(image_path):
    with open(image_path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")

def extract_text_from_mp4(file_path,key):
    client = OpenAI(api_key= "") #add your api
    audio_file= open(file_path, "rb")
    transcription = client.audio.transcriptions.create(model="gpt-4o-transcribe", file=audio_file)
    return transcription.text
    
def extract_text_from_code_file(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            return file.read()
    except UnicodeDecodeError:
        try:
            with open(filename, 'r', encoding='latin-1') as file:
                return file.read()
        except Exception as e:
            logging.error(f"Error reading {filename}: {e}")
            return ""


def extract_text_from_text_file(filename):
    try:
        with open(filename, 'r', encoding='utf-8') as file:
            return file.read()
    except UnicodeDecodeError:
        with open(filename, 'r', encoding='latin1') as file:
            return file.read()
        
def extract_text_from_excel(filename):
    try:
        df = pd.read_excel(filename)
        return df.to_string(index=False)
    except Exception as e:
        logging.error(f"Error reading {filename}: {e}")
        return ""

def extract_text_from_doc(filename):
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(filename)
        text = doc.Content.Text
        doc.Close()
        word.Quit()
        return text
    except Exception as e:
        logging.error(f"Error reading {filename}: {e}")
        return ""

def extract_text_from_file(filename):
    if filename.endswith('.pdf'):
        return extract_text_from_pdf(filename)
    elif filename.endswith('.docx'):
        return extract_text_from_docx(filename)
    elif filename.endswith('.doc'):
        return extract_text_from_doc(filename)
    elif filename.endswith(('.png', '.jpg', '.jpeg')):
        return extract_text_from_image(filename)
    elif filename.endswith(('.c', '.h')):
        return extract_text_from_code_file(filename)
    elif filename.endswith('.txt'):
        return extract_text_from_text_file(filename)
    elif filename.endswith('.xls') or filename.endswith('.xlsx'):
        return extract_text_from_excel(filename)
    else:
        logging.warning(f"Unsupported file format: {filename}")
        return ""

def list_files_in_directory_by_type(directory):
    pdfs = []
    docs = []
    images = []
    codes = []
    texts = []
    excels = []
    audio = []

    for file in os.listdir(directory):
        file_path = os.path.join(directory, file)
        if os.path.isfile(file_path):
            if file.endswith('.pdf'):
                pdfs.append(file_path)
            elif file.endswith('.docx') or file.endswith('.doc'):
                docs.append(file_path)
            elif file.endswith(('.png', '.jpg', '.jpeg')):
                images.append(file_path)
            elif file.endswith(('.c', '.h')):
                codes.append(file_path)
            elif file.endswith('.txt'):
                texts.append(file_path)
            elif file.endswith('.mp4'):
                audio.append(file_path)
            elif file.endswith('.xls') or file.endswith('.xlsx'):
                excels.append(file_path)

    return pdfs, docs, images, codes, texts, excels, audio

# Add system instructions
def mathematics_competitions(models,think):
    setup_logging(models) 
    logging.info(f'Starting evaluation using model: {models}')
    
    system_instructions = "你是一名具有专业背景的STEM教育专家，需要对学生提交的项目式学习成果进行评估。"

    query = """你是一名具有专业背景的STEM教育专家，需要对学生提交的多模态项目式学习成果进行评估，根据学生提交的文档、代码、图片、音频等材料，评估其在每个子指标上的表现（以文档为主）。多模态的项目式学习成果的总评分采用四个核心维度的加权综合方式进行计算，满分为100分。
    1. 知识性维度包含： 1） 概念理解（15分） 2） 跨学科应用（5分） 3）证据推理能力（5分） 
    2. 技能型维度包含：1） 工具与流程应用（10分） 2）问题解决能力（10分） 3）自主调控与计划（10分）
    3. 表达型维度包含：1）信息表达清晰度（8分） 2）多模态表达能力（6分） 3）受众意识与适应（6分） 
    4. 创新反思维度包含：1）创新性与实用性（15分） 2）创意发展与迭代（5分） 3）自我反思与成长（5分）
    你的任务：请针对上述12个子指标，根据学生提交的多模态材料，逐项提供：每个子指标对应的评分，并一句简要评分理由（可依据内容质量、逻辑、完成度等）。最后给出整个项目的总分（仅输出一个整数）。"""
    
    main_directory = "./"
    competition_dirs = [os.path.join(main_directory, d) for d in os.listdir(main_directory) if os.path.isdir(os.path.join(main_directory, d))]
   
    for competition_dir in competition_dirs:
        project_dirs = [os.path.join(competition_dir, d) for d in os.listdir(competition_dir) if os.path.isdir(os.path.join(competition_dir, d)) and d.startswith("project")]
        for project_dir in project_dirs:
            keys = [""]#add your api
            key = random.choice(keys)
            client = genai.Client(api_key=key)

            
            print(project_dir)
            messages = []
            messages.append(system_instructions)
            
            content = {"role": "user", "parts": ""}
            pdfs, docs, images, codes, texts, excels, audio = list_files_in_directory_by_type(project_dir)

            if pdfs or docs:
                content["parts"] += "以下是项目文档：\n"
                for file in pdfs + docs:
                    text = extract_text_from_file(file)
                    content["parts"] += text + "\n"

            if images:
                content["parts"] += "以下是项目图片：\n"
                for file in images:
                    image_file = client.files.upload(file=file)
                    messages.append(image_file)

            if codes:
                content["parts"] += "以下是项目代码：\n"
                for file in codes:
                    text = extract_text_from_file(file)
                    content["parts"] += text + "\n"
                    
            if audio:
                content["parts"] += "以下是项目视频描述：\n"
                for file in audio:
                    text = extract_text_from_mp4(file,key)
                    content["parts"] += text + "\n"
                    
            if texts or excels:
                content["parts"] += "以下是其他项目内容：\n"
                for file in texts + excels:
                    text = extract_text_from_file(file)
                    content["parts"] += text + "\n" 
    
            messages.append(content["parts"])
            messages.append(query)
            

            
            scores = []
            for i in range(5):
                try:
                    if think:
                        completion = client.models.generate_content(
                                                    model=models,
                                                    contents=messages,
                                                    config=types.GenerateContentConfig(thinking_config=types.ThinkingConfig(thinking_budget=256)))
                    else:
                        completion = client.models.generate_content(
                                                    model=models,
                                                    contents=messages)
                    score = extract_answer(completion.text)
                    logging.info(f"{project_dir}: 第 {i+1} 次评估评分: {score}")
                    print(f"第 {i+1} 次评估评分：{score}")
                    scores.append(score)
                    time.sleep(80)
                except Exception as e:
                    logging.error(f"Error processing {project_dir}: {e}")
                    
            mean = calculate_mean(scores)
            std_dev = calculate_std_dev(scores)
            print(f"平均值: {mean}")
            print(f"标准差: {std_dev}")
            logging.info(f"{project_dir}: 平均值: {mean}")
            logging.info(f"{project_dir}: 平均值: {std_dev}")   
