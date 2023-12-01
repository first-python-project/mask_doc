# 한글 파일인 경우 특정 단어가 있으면 마스킹하는 함수
# txt파일을 통해 특정단어 설정
# 기존 파일 수정

from docx import Document
import os

#dir_path = os.getcwd()
dir_path = 'static'
masking_word_txt = 'word.txt'

def process_document(dir_path, masking_word_txt):
    # txt파일에 저장된 특정 단어 읽기
    def read_masking_word(file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            # 공백제거
            return [line.strip() for line in lines]
        
    masking_word = read_masking_word(masking_word_txt)

    # 경로의 doc, docx 파일 찾기 
    for file in os.listdir(dir_path):
        if file.endswith(".docx") or file.endswith(".doc"):
            input_doc = os.path.join(dir_path, file)
            output_doc = os.path.join(dir_path, f"{os.path.splitext(file)[0]}.docx")

            doc = Document(input_doc)

            # doc, docx 파일 특정 단어 변경
            for file_doc in doc.paragraphs:
                for word in masking_word:
                    if word in file_doc.text:
                        # 단어길이만큼 *로 변경
                        masking = '*' * len(word)
                        file_doc.text = file_doc.text.replace(word, masking)
                        
            doc.save(output_doc)

            

process_document(dir_path,masking_word_txt)