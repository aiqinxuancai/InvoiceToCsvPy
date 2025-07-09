import os
import sys
import csv
import json
import time # 导入time模块用于延时
import re   # 导入re模块用于处理非法字符
from pathlib import Path
from openai import OpenAI

# --- 配置区域 ---
API_KEY_FILE = "moonshot.txt"
PDF_FOLDER_PATH = "./invoices"
CSV_OUTPUT_PATH = "invoices_data.csv"
MAX_RETRIES = 3           # API请求最大重试次数
RETRY_DELAY = 2           # 每次重试的间隔时间（秒）
# --- 配置区域结束 ---

# 将CSV表头定义为全局常量，方便多处使用
CSV_HEADER = [
    "发票代码", "发票号码", "销方识别号", "销方名称", "购方识别号", "购买方名称",
    "开票日期", "项目名称", "数量", "金额", "税率", "税额", "价税合计",
    "发票票种", "类别"
]

client = None

def get_api_key(filename):
    """从文件中读取API密钥。"""
    try:
        if not os.path.exists(filename):
             raise FileNotFoundError
        with open(filename, 'r', encoding='utf-8') as f:
            # .strip() 用于移除可能存在的多余空格或换行符
            key = f.read().strip()
            if not key:
                print(f"错误：API密钥文件 '{filename}' 是空的。")
                input("请在文件中填入您的Kimi API密钥后重试。按 Enter 键退出。")
                sys.exit(1)
            return key
    except FileNotFoundError:
        print(f"错误：未找到API密钥文件 '{filename}'。")
        print(f"请在程序目录下创建一个名为 '{filename}' 的文本文件，并将您的Kimi API密钥粘贴进去。")
        input("按 Enter 键退出。") # 暂停程序，方便用户看到错误信息
        sys.exit(1)

def sanitize_filename(name):
    """移除文件名中的非法字符。"""
    return re.sub(r'[\\/*?:"<>|]', "", name)

def create_na_record():
    """创建一个所有字段均为'N/A'的记录。"""
    return {header: "N/A" for header in CSV_HEADER}

def rename_successful_pdf(original_path, data):
    """根据提取的数据重命名PDF文件。"""
    try:
        invoice_num = data.get("发票号码", "N_A")
        issue_date = data.get("开票日期", "N_A")
        category = data.get("类别", "N_A")
        total_amount_str = str(data.get("价税合计", "0"))

        # 将总金额转换为不带小数点的整数
        try:
            total_amount_int = int(float(total_amount_str))
        except (ValueError, TypeError):
            total_amount_int = 0

        # 构建新文件名并清理非法字符
        new_base_name = f"{invoice_num}-{issue_date}-{category}-{total_amount_int}.pdf"
        sanitized_name = sanitize_filename(new_base_name)
        
        dir_name = os.path.dirname(original_path)
        new_path = os.path.join(dir_name, sanitized_name)

        # 检查新文件名是否已存在，避免覆盖
        if os.path.exists(new_path):
            print(f"警告：重命名失败，文件 '{new_path}' 已存在。")
            return

        os.rename(original_path, new_path)
        print(f"文件已成功重命名为: {sanitized_name}")

    except Exception as e:
        print(f"警告：文件 '{os.path.basename(original_path)}' 重命名失败: {e}")

def extract_invoice_info_from_pdf(pdf_path):
    """上传PDF，提取信息，包含重试和重命名逻辑。"""
    if not client:
        print("错误：API客户端未初始化。")
        return create_na_record()
        
    print(f"--- 开始处理文件: {os.path.basename(pdf_path)} ---")
    
    # --- API请求与重试逻辑 ---
    completion = None
    file_object = None
    for attempt in range(MAX_RETRIES):
        try:
            # 仅在第一次尝试时上传文件
            if not file_object:
                file_object = client.files.create(file=Path(pdf_path), purpose="file-extract")
            
            file_content = client.files.content(file_id=file_object.id).text
            
            # 构建完整的提取信息Prompt
            prompt = f"""
            你是一个发票信息提取助手。请从下面的文本内容中提取结构化的发票信息，其中发票票种为"增值税电子普通发票"或"电子发票（普通发票）"。

            文件内容:
            ---
            {file_content}
            ---

            请严格按照以下JSON格式返回提取的信息，如果某个字段在文件中不存在，请用 "N/A" 表示。
            {{
                "发票代码": "invoice_code",
                "发票号码": "invoice_number",
                "销方识别号": "seller_tax_id",
                "销方名称": "seller_name",
                "购方识别号": "buyer_tax_id",
                "购买方名称": "buyer_name",
                "开票日期": "issue_date",
                "项目名称": "item_name",
                "数量": "quantity",
                "金额": "amount",
                "税率": "tax_rate",
                "税额": "tax_amount",
                "价税合计": "total_amount",
                "发票票种": "invoice_type",
                "类别": "category"
            }}

            提取的例子：
            ```csv
            发票代码,发票号码,销方识别号,销方名称,购方识别号,购买方名称,开票日期,项目名称,数量,金额,税率,税额,价税合计,发票票种,类别
            ,24442000000657111111,91440812MAD2B8BJX7,湛江市西海岸西厨餐饮管理有限公司,91111111MA9W511111,广州咖喱网络科技有限公司,2024年12月30日,*餐饮服务*餐饮服务,1,109.9,1%,1.1,111,电子发票（普通发票）,餐饮服务
            044002301111,45311111,9144000061740323XQ,百胜餐饮（广东）有限公司,91111111MA9W511111,广州咖喱网络科技有限公司,2024年12月30日,*餐饮服务*餐饮服务,1,177.28,6%,10.64,187.92,增值税电子普通发票,餐饮服务
            ```

            类别请在下面类目中匹配：
            [餐饮服务,住宿服务,交通运输服务,居民日常服务,办公用品,电子设备,咨询服务,技术服务,租赁服务,建筑服务,医疗服务,教育服务,商品零售]

            金额相关数值请不要带有货币符号，例如¥等。
            """

            messages = [
                {
                    "role": "system",
                    "content": "你是一个专业的发票数据提取机器人，请根据用户提供的发票文本内容，准确地抽取出关键信息并以JSON格式返回。"
                },
                {
                    "role": "user",
                    "content": prompt,
                }
            ]

            completion = client.chat.completions.create(
                model="moonshot-v1-32k",
                messages=messages,
                temperature=0.0,
                response_format={"type": "json_object"}
            )
            print(f"API请求成功 (尝试 {attempt + 1}/{MAX_RETRIES})")
            break  # 成功则跳出重试循环

        except Exception as e:
            print(f"API请求失败 (尝试 {attempt + 1}/{MAX_RETRIES}): {e}")
            if attempt < MAX_RETRIES - 1:
                print(f"将在 {RETRY_DELAY} 秒后重试...")
                time.sleep(RETRY_DELAY)
            else:
                print("所有重试均失败。")
    
    # 清理已上传的文件
    if file_object:
        try:
            client.files.delete(file_id=file_object.id)
        except Exception as e:
            print(f"警告：清理云端文件 {file_object.id} 失败: {e}")

    # --- 处理API结果 ---
    if not completion:
        print(f"处理失败，为文件 {os.path.basename(pdf_path)} 插入N/A记录。")
        return create_na_record()

    try:
        extracted_data = json.loads(completion.choices[0].message.content)
        print("API响应解析成功。")
        
        # 成功提取数据后，重命名文件
        rename_successful_pdf(pdf_path, extracted_data)
        
        return extracted_data
    except Exception as e:
        print(f"解析API响应或重命名文件时出错: {e}")
        return create_na_record()


def main():
    """主函数，初始化并执行所有流程。"""
    global client
    
    api_key = get_api_key(API_KEY_FILE)
    client = OpenAI(api_key=api_key, base_url="https://api.moonshot.cn/v1")

    if not os.path.isdir(PDF_FOLDER_PATH):
        print(f"错误：文件夹 '{PDF_FOLDER_PATH}' 不存在。")
        input("按 Enter 键退出。")
        return

    pdf_files = [os.path.join(PDF_FOLDER_PATH, f) for f in os.listdir(PDF_FOLDER_PATH) if f.lower().endswith('.pdf')]
    if not pdf_files:
        print(f"在文件夹 '{PDF_FOLDER_PATH}' 中没有找到PDF文件。")
        input("按 Enter 键退出。")
        return

    print(f"找到 {len(pdf_files)} 个PDF文件。开始处理...")

    all_data = []
    for pdf_path in pdf_files:
        # 检查文件是否还存在，因为它可能已被成功处理并重命名
        if os.path.exists(pdf_path):
            invoice_data = extract_invoice_info_from_pdf(pdf_path)
            all_data.append(invoice_data)
        else:
            print(f"跳过文件 {os.path.basename(pdf_path)}，因为它可能已被重命名。")


    with open(CSV_OUTPUT_PATH, mode='w', newline='', encoding='utf-8-sig') as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=CSV_HEADER)
        writer.writeheader()
        writer.writerows(all_data)

    print(f"\n--- 处理完成！---\n所有发票信息已保存到文件: {CSV_OUTPUT_PATH}")
    input("按 Enter 键退出。")

if __name__ == '__main__':
    main()