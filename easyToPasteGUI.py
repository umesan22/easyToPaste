from PIL import Image, ImageGrab, ImageTk
from tkinter import messagebox, filedialog
from time import sleep
import os, sys, tkinter
import win32clipboard as wc
import xlwings as xl

def file_select():
    # 選択可能なファイル形式
    fTyp = [('','*.xlsx')]
    # ファイル選択エクスプローラー初期画面
    iDir = os.path.abspath(os.path.dirname(__file__))
    # 選択したファイルパス
    file_path = filedialog.askopenfilename(filetypes = fTyp, initialdir = iDir)
    file_edit.insert(tkinter.END, file_path)

def validate():
    is_validate_ok = False
    if not top_edit.get().isdecimal() or int(top_edit.get()) < 1:
        is_validate_ok = True
    if not left_edit.get().isdecimal() or int(left_edit.get()) < 1:
        is_validate_ok = True
    if not space_edit.get().isdecimal() or int(space_edit.get()) < 1:
        is_validate_ok = True
    
    if not is_new_excel.get():
        if len(file_edit.get()) < 1:
            messagebox.showwarning('File Not Found', 'Excelファイルを指定してください。')
    if not is_resize_pic.get():
        if not width_edit.get().isdecimal() or int(width_edit.get()) < 1:
            is_validate_ok = True
        if not height_edit.get().isdecimal() or int(height_edit.get()) < 1:
            is_validate_ok = True

    if is_validate_ok:
        messagebox.showwarning('Validation Error', '''\
            Excel,シート選択項目以外の
            各項目は1以上の整数で
            入力してください。
            ''')

    return is_validate_ok

def invisible():
    if is_new_excel.get():
        # ファイル選択を隠す
        file_edit.configure(state='readonly')
        sheet_edit.configure(state='readonly')
    else:
        # 出現させる
        file_edit.configure(state='normal')
        sheet_edit.configure(state='normal')

    if is_resize_pic.get():
        width_edit.configure(state='readonly')
        height_edit.configure(state='readonly')
    else:
        width_edit.configure(state='normal')
        height_edit.configure(state='normal')

def pic_paste():
    is_validate_ok = validate()
    if is_validate_ok:
        return

    # エクセルファイルを取得
    file_path = file_edit.get()
    sheet = sheet_edit.get()

    # 貼り付け開始位置
    top_px = int(top_edit.get())
    left_px = int(left_edit.get())
    space = int(space_edit.get())

    # 画像のリサイズ
    if not is_resize_pic.get():
        pic_width = int(width_edit.get())
        pic_height = int(height_edit.get())
    pic_path = os.getcwd() + '\\ScreenShot.png'

    data = None
    is_exist = False

    # 貼り付けるエクセルを開く
    try:
        if is_new_excel.get():
            book = xl.App(visible=None, add_book=False).books.add()
        else:
            book = xl.App(visible=None, add_book=False).books.open(file_path)
            for names in book.sheets:
                if names.name == sheet:
                    is_exist = True
                elif len(sheet) > 0:
                    messagebox.showwarning('Sheet Not Found', text=sheet + 'はありませんでした。')
                    return
    except:
        messagebox.showwarning('File Not Found', 'ファイルが見つかりませんでした。')
        return

    if not is_new_excel.get() and is_exist:
        book.sheets(sheet).activate()

    # 2秒ごとにクリップボードを監視
    while True:
        try:
            wc.OpenClipboard()

            if not wc.IsClipboardFormatAvailable(wc.CF_DIB):
                wc.CloseClipboard()
                sleep(2)
                continue

            img = wc.GetClipboardData(wc.CF_DIB)

            if data is not None:
                # クリップボードの画像と同じならコンティニュー
                if data == img:
                    wc.CloseClipboard()
                    sleep(2)
                    continue
            
            im = ImageGrab.grabclipboard()
            if isinstance(im, Image.Image):
                if not is_resize_pic.get():
                    # 画像をリサイズ
                    resize_im = im.resize((pic_width, pic_height), Image.ANTIALIAS)
                    # 一度画像を保存
                    resize_im.save(pic_path, quality=100)
                else:
                    im.save(pic_path, quality=100)
                
                # 写真をエクセルに貼り付け
                book.app.range('A1').sheet.pictures.add(pic_path, top=top_px, left=left_px)
                # 画像を削除
                os.remove(pic_path)
                data = img
                if not is_resize_pic.get():
                    top_px += (pic_height + space)
                else:
                    top_px += (int(str(im.size).split(',')[1].strip().rstrip(')')) + space)
            
            sleep(2)

        except:
            messagebox.showerror('ERROR', '''\
                例外が発生しました。
                アプリケーションを終了します。
                ''')
            sys.exit()

# GUI作成
root = tkinter.Tk()
root.title('easyToPaste')
root.geometry('500x225')

# 必要なラベルと入力フォームを作成
top_label = tkinter.Label(text='上貼り付け開始位置(px)')
top_label.place(x=15, y=15)
top_edit = tkinter.Entry()
top_edit.place(x=15, y=35)

left_label = tkinter.Label(text='左貼り付け開始位置(px)')
left_label.place(x=175, y=15)
left_edit = tkinter.Entry()
left_edit.place(x=175, y=35)

space_label = tkinter.Label(text='貼り付け画像の縦の間隔(px)')
space_label.place(x=335, y=15)
space_edit = tkinter.Entry()
space_edit.place(x=335, y=35)

width_label = tkinter.Label(text='貼り付け画像幅(px)')
width_label.place(x=15, y=65)
width_edit = tkinter.Entry()
width_edit.place(x=15, y=85)

height_label = tkinter.Label(text='貼り付け画像高さ(px)')
height_label.place(x=175, y=65)
height_edit = tkinter.Entry()
height_edit.place(x=175, y=85)

is_resize_pic = tkinter.BooleanVar()
is_resize_pic.set(False)
check_original = tkinter.Checkbutton(text='画像サイズを変えない', variable=is_resize_pic, command=invisible)
check_original.place(x=315, y=85, height=20)

file_label = tkinter.Label(text='貼り付けするExcelを選択')
file_label.place(x=15, y=115)
file_edit = tkinter.Entry(width=40)
file_edit.place(x=15, y=135)
file_button = tkinter.Button(text='参照', command=file_select)
file_button.place(x=270, y=135, height=20, width=50)

is_new_excel = tkinter.BooleanVar()
is_new_excel.set(False)
check_new_file = tkinter.Checkbutton(text='新規ファイル', variable=is_new_excel, command=invisible)
check_new_file.place(x=330, y=135, height=20, width=100)

sheet_label = tkinter.Label(text='貼り付けするシートを選択')
sheet_label.place(x=15, y=165)
sheet_edit = tkinter.Entry(width=40)
sheet_edit.place(x=15, y=185)

start_button = tkinter.Button(text='開始', bg='#00FF00', command=pic_paste)
start_button.place(x=335, y=185, height=20, width=80)

root.mainloop()
