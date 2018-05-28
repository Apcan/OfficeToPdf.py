import sys
import os
import win32com.client

wd_pdf_format = 17
pt_pdf_format = 32
xl_pdf_format = 0

wdformat = ['.doc', '.docx']
ptformat = ['.ppt', '.pptx']
xlformat = ['.xls', '.xlsx']

argc = len(sys.argv)

if argc < 2:
    sys.exit(-1)


src = sys.argv[1]
dst = sys.argv[2]

if not os.path.exists(src):
    print('文件不存在')
    sys.exit(-1)

ext = os.path.splitext(src)[1]


def wordtopdf(src, dst):
    os.system('taskkill /f /im WINWORD.EXE')
    w = win32com.client.Dispatch('Word.Application')
    w.Visible = 0
    w.DisplayAlerts = 0
    worddoc = w.Documents.Open(src)
    worddoc.SaveAs(dst, wd_pdf_format)
    worddoc.Close()
    w.Quit()
    sys.exit(0)
    return


def ppttopdf(src, dst):
    os.system('taskkill /im POWERPNT.EXE')
    p = win32com.client.Dispatch('PowerPoint.Application')
    p.Visible = 1
    p.DisplayAlerts = 0
    ppt = p.Presentations.Open(src)
    ppt.SaveAs(dst, pt_pdf_format)
    ppt.Close()
    p.Quit()
    sys.exit(0)
    return


def xlstopdf(src, dst):
    os.system('taskkill /im EXCEL.EXE')
    x = win32com.client.Dispatch('Excel.Application')
    x.Visible = 0
    x.DisplayAlerts = 0
    xls = x.Workbooks.Open(src)
    xls.ExportAsFixedFormat(xl_pdf_format, dst)
    xls.Close()
    x.Quit()
    sys.exit(0)
    return


if (ext in wdformat):
    wordtopdf(src, dst)
elif (ext in ptformat):
    ppttopdf(src, dst)
elif (ext in xlformat):
    xlstopdf(src, dst)
