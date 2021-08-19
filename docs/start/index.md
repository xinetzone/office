# 快速上手

告诉读者如何做某事的简短指南。假定读者拥有一些背景知识（通常，读者已经过阅读教程）。读者专注于“做事”而不是长篇大论的解释。How To 指南构成了大部分的文档，并且针对不同的主题分为不同的部分。

## Excel

- [xlwings](https://www.xlwings.org/) 不光可以读写 excel，还能进行格式调整、VBA 操作，非常强大且易于使用。
- pandas
- [Python Resources for working with Excel - Working with Excel Files in Python (python-excel.org)](http://www.python-excel.org/)

## PPT

[python-pptx](https://python-pptx.readthedocs.io/en/latest/) 支持 ppt 的自动化处理，主要的库还有 pywin32com、pptx，可以创建、修改 ppt 文件。

## Word

* **[python-docx](https://python-docx.readthedocs.io/en/latest/)、import docx** ：只对 windows 平台有效
* **pypiwin32、import win32com** ：跨平台，但无法处理 doc 格式的 word 文本，doc 格式不是基于 xml 的
* **textract、import textract** ：它同时兼顾“doc”和“docx”，但安装过程需要一些依赖。 你可以批量的用 python 生成 word 文件，推荐使用 docx，不需要会太多。

## **邮件处理**

python 处理邮件也是极其便利的，smtplib、imaplib、email 三个库配合使用，实现邮件编写、发送、接收、读取等一系列自动化操作，省时省力。

## PDF

[ReportLab - Content to PDF Solutions](https://www.reportlab.com/)
