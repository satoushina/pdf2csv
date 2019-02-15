# encoding: utf-8 CRLF

#
#http://irukanobox.blogspot.com/2017/03/python3pdf.html

#厚生労働省のブラック企業リストをPythonで解析する（PDFMiner.six）
#   https://a244.hateblo.jp/entry/2017/08/26/001050

#Programming with PDFMiner
#   https://euske.github.io/pdfminer/programming.html

#pdfをリスト(csv)にする試み



#コマンド読み込み　一般
import sys
import math
import re
#from operator import attrgetter


#必要コマンド読み込み mdfminer.six
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice

from pdfminer.converter import TextConverter, PDFPageAggregator
from pdfminer.layout import LAParams

from pdfminer.layout import LTTextContainer, LTTextBox

from io import StringIO

DEBUG = True
#DEBUG = False



# モジュールのインポート
import os, tkinter, tkinter.filedialog, tkinter.messagebox
#ファイル入力 #コマンドライン引数をfilenameに渡す
if DEBUG:
    print(sys.argv)

if ( len(sys.argv) > 1 ):
    filepath = sys.argv[1]
    if DEBUG:
        print(filepath)
else:
    # ファイル選択ダイアログの表示
    # https://qiita.com/chanmaru/items/1b64aa91dcd45ad91540
    root = tkinter.Tk()
    root.withdraw()
    fTyp = [("","*")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    tkinter.messagebox.showinfo('pdfからcsv変換プログラム','処理ファイルを選択してください！')
    filepath = tkinter.filedialog.askopenfilename(filetypes = fTyp,initialdir = iDir)

    # 処理ファイル名の出力
    if DEBUG:
        tkinter.messagebox.showinfo('pdfからcsv変換プログラム',filepath)

    #exit(-1)


#https://note.nkmk.me/python-pathlib-name-suffix-parent/
import pathlib
InputFile = pathlib.Path(filepath)
if DEBUG:
    print(type(InputFile))
filebasename = InputFile.name
filebasestem = InputFile.stem
filedir = InputFile.parent


#import os
#filename = os.path.basename(filepath)
#filedir  = os.path.dirname(filepath)
#sepstr = os.sep






# 処理するPDFを開く # Open a PDF file.
fp = open(filepath, 'rb')
#パーサーをセット？ # Create a PDF parser object associated with the file object.
parser = PDFParser(fp)
# Create a PDF document object that stores the document structure.
# Supply the password for initialization.
document = PDFDocument(parser, password='')
# Check if the document allows text extraction. If not, abort.
if not document.is_extractable:
    raise PDFTextExtractionNotAllowed
#リソースマネージャーをセット？
rsrcmgr = PDFResourceManager()




#LAパラメーターをセット？
laparams = LAParams()
# 縦書き文字があれば、横並びで出力する
laparams.detect_vertical = True
#rettxt = StringIO()
device = PDFPageAggregator(rsrcmgr, laparams=laparams)
#機能の設定？
#device = TextConverter(rsrcmgr, rettxt, codec='utf-8', laparams=laparams)
#インタープリターをセット？
interpreter = PDFPageInterpreter(rsrcmgr, device)




for page in PDFPage.create_pages(document):
    interpreter.process_page(page)
    # receive the LTPage object for the page.
    layout = device.get_result()
    if DEBUG :
        print("pageid = " + str(layout.pageid) )



    cells = list()

    for node in layout:
        if ( not issubclass(node.__class__ ,(LTTextBox, LTTextContainer ) ) ):
            continue
        else:
            cells.append(node)
            temp_str = node.get_text().rstrip("\n")
            if DEBUG :
                #print(temp_str)
                continue



#cellsリストが完成
#cellsから内容を削ぎ落として出力したい表に変えるプロセスを以下に書く
if DEBUG :
    print("リストcellsの要素数は" + str( len(cells) ) )


#cellsの中身でほしいものをそれぞれリストに入れ直す
#array = [cells.x0 , cells.y0 , cells]
#celx = list()
#cely = list()
#celt = list()
array = list()
for cell in cells:
    tempcell = list()
    #celx.append(cell.x0) #左座標点
    #cely.append(cell.y0) #下座標点 #y1は上座標点
    #celt.append(cell.get_text().rstrip("\n") ) #末尾の改行を消す
    tempcell = [cell.x0 , cell.y1 , cell.get_text().rstrip("\n")]
    array.append(tempcell)
#array = [celx , cely, celt]


if DEBUG :
    print("リストarrayの要素数は" + str( len(array) ) )


#並び替えたい
#https://qiita.com/fantm21/items/6df776d99356ef6d14d4
#data.sort(key = lambda x:x[0] , reverse=True)
#data.sort(key = lambda x:x[1] )
#いまいち　列１が第一優先　昇順　列２が第二優先　降順

#並び替え　こちらを採用かな？
#https://docs.python.jp/3/howto/sorting.html
from operator import itemgetter, attrgetter
array.sort(key = itemgetter(0))
array.sort(key = itemgetter(1),reverse=True)




#行列を入れ替える
#import numpy as np
#data = np.array(array).T.tolist()
#pdfは左下原点の座標系 第一象限
#data = array




#csv出力する
import csv
if DEBUG:
    print(InputFile.with_suffix(".csv"))
with open(InputFile.with_suffix(".csv") , "w") as file:
    writer = csv.writer(file, lineterminator='\n')
    writer.writerows(array)















fp.close()
device.close()
#rettxt.close()