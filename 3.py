import importlib
import sys
import time
import os
import re

from PyPDF2 import PdfReader #pdf的读取方法
from PyPDF2 import PdfWriter #pdf的写入方法
"""
如果是python2的将上面的PdfReader和PdfWriter改为
PdfFileReader和PDFfileWriter即可
"""
from Crypto.Cipher import AES #高加密的方法，要引入不然会报错

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfpage import PDFTextExtractionNotAllowed

importlib.reload(sys)
time1 = time.time()

def get_reader(filename, password):
    try:
        old_file = open(filename, 'rb')
        print('解密开始...')
    except Exception as err:
        return print('文件打开失败！' + str(err))

    pdf_reader = PdfReader(old_file)

    if pdf_reader.is_encrypted:
        if password is None:
            return print('文件被加密，需要密码！--{}'.format(filename))
        #else:
            #if pdf_reader.decrypt(password) == 1:
                #return print('密码不正确！--{}'.format(filename))

    if old_file in locals():
        old_file.close()

    return pdf_reader

def deception_pdf(filename, password, decrypted_filename=None):
    print('正在生成解密...')
    pdf_reader = get_reader(filename, password)
    if pdf_reader is None:
        return print("无内容读取")

    if not pdf_reader.is_encrypted:
        return print('文件没有被加密，无需操作')

    pdf_writer = PdfWriter()

    for page in pdf_reader.pages:
        pdf_writer.add_page(page)

    if decrypted_filename is None:
        decrypted_filename = "".join(filename.split('.')[:-1]) +'.pdf'
        print("解密文件已生成:{}".format(decrypted_filename))

    with open(decrypted_filename, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)

def parse(pdf_path, txt_path):
    # 解析PDF文本，并保存到TXT文件中
    fp = open(pdf_path, 'rb')
    # 用文件对象创建一个PDF文档分析器
    parser = PDFParser(fp)
    # 创建一个PDF文档
    doc = PDFDocument(parser)
    # 连接分析器，与文档对象
    parser.set_document(doc)
    #doc.set_parser(parser)

    # 提供初始化密码，如果没有密码，就创建一个空的字符串
    #doc.initialize()

    # 检测文档是否提供txt转换，不提供就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        # 创建PDF，资源管理器，来共享资源
        rsrcmgr = PDFResourceManager()
        # 创建一个PDF设备对象
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        # 创建一个PDF解释器对象
        interpreter = PDFPageInterpreter(rsrcmgr, device)

        # 循环遍历列表，每次处理一个page内容
        # doc.get_pages() 获取page列表
        for i, page in enumerate(PDFPage.create_pages(doc)):
            interpreter.process_page(page)
            # 接受该页面的LTPage对象
            layout = device.get_result()
            # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象
            # 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等
            # 想要获取文本就获得对象的text属性，
            for x in layout:
                if (isinstance(x, LTTextBoxHorizontal)):
                    with open(txt_path, 'a+', encoding='utf-8') as f:
                        results = x.get_text()
                        print(results)
                        f.write(results + "\n")

def process_files(folder_path):

   # 获取文件总数
   total_files = len([filename for filename in os.listdir(folder_path) if filename.endswith('.pdf')])
   processed_files = 0

   # 遍历文件夹中的所有pdf文件
   for filename in os.listdir(folder_path):
       if filename.endswith('.pdf'):
           # 解析文件名，提取股票代码、公司简称和年份

           match = re.match(r'^(\d{6})_(.*?)_(\d{4})\.pdf$', filename)
           #if match:
               #stock_code = match.group(1)
               #company_name = match.group(2)
               #year = match.group(3)
           pdf_path = 'C://Users\末路歧途\Desktop\LC 爬取\年报PDF版\\'+filename

           deception_pdf(pdf_path, '')#解密PDF

           txt_path = 'C://Users\末路歧途\Desktop\LC 爬取\年报TXT版\\'+match.group(1)+'_'+match.group(2)+'_'+match.group(3)+'.txt'
           parse(pdf_path, txt_path)
           time2 = time.time()
           print("总共消耗时间为:", time2 - time1)

if __name__ == '__main__':
    process_files('C://Users\末路歧途\Desktop\LC 爬取\年报PDF版')

