from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from docxcompose.composer import Composer
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_BREAK
from docx.oxml.ns import qn

"""from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, ToolMessage
import yfinance as yf
from langchain.tools import tool
from langchain_core.prompts import PromptTemplate
import pandas as pd
import numpy as np
import os
import requests
import re
from datetime import datetime
import json
from langchain_core.output_parsers import CommaSeparatedListOutputParser
from serpapi import GoogleSearch"""

"""template=DocxTemplate('test_docx.docx')

etf_name='qqq'
etf_detail="기술 지수"

context={
    'etf_name': etf_name,
    'etf_detail':etf_detail
}

template.render(context)

num=3
title1='result'
title2='.docx'

final_title=title1+str(num)+title2

template.save(final_title)"""

def merge_docs(output_path, *input_paths):
    # 첫 번째 문서를 기본 문서로 사용
    base_doc = Document(input_paths[0])
    composer = Composer(base_doc)

    # 나머지 문서들을 순회하며 병합
    for file_path in input_paths[1:]:
        doc = Document(file_path)
        composer.append(doc)

    # 병합된 문서 저장
    composer.save(output_path)
    print(f"문서들이 성공적으로 병합되어 {output_path}에 저장되었습니다.")

if __name__ == "__main__":
    output_file = "result_final.docx"
    input_files = ["result1.docx", "result2.docx", "result3.docx"]
    merge_docs(output_file, *input_files)