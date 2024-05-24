from docx import Document
import re
import codecs

def validate_id(id_number):
    # 驗證身分證字號格式
    pattern = re.compile(r'^[A-Z][1-2]\d{8}$')
    if not pattern.match(id_number):
        return False
    
    # 計算檢查碼
    letters = 'ABCDEFGHJKLMNPQRSTUVXYWZIO'
    id_letters = letters.index(id_number[0]) + 10
    total = id_letters // 10 + (id_letters % 10) * 9
    for i, digit in enumerate(id_number[1:9]):
        total += int(digit) * (8 - i)
    return str(10 - total % 10) == id_number[-1]

def extract_valid_ids(file_path):
    document = Document(file_path)
    
    valid_ids = []
    for paragraph in document.paragraphs:
        matches = re.findall(r'\b[A-Z][1-2]\d{8}\b', paragraph.text)
        for match in matches:
            if validate_id(match):
                valid_ids.append(match)
    return valid_ids

def export_to_txt(ids, output_file):
    # 匯出符合身分證格式的文字符到新的UTF-8編碼的文本文件
    with codecs.open(output_file, 'w', 'utf-8') as f:
        f.write('符合身分證格式的文字符：\n\n')
        for id_number in ids:
            f.write(id_number + '\n\n')

if __name__ == "__main__":
    # 輸入Word文件路徑
    input_file_path = './test2.docx'
    # 分析Word文件，找出符合身分證格式的文字符
    print("身分證字號-掃描結果：")
    valid_ids = extract_valid_ids(input_file_path)
    for id_number in valid_ids:
        print(id_number)
    # 匯出符合身分證格式的文字符到新的UTF-8編碼的文本文件
    export_to_txt(valid_ids, 'output.txt')

    printf()
