
#参考：https://qiita.com/Tomo666/items/1b64aa91dcd45ad91540

# モジュールのインポート
import os, tkinter, tkinter.filedialog, tkinter.messagebox

def py_gui_dir():
    # フォルダ選択ダイアログの表示
    root = tkinter.Tk()
    root.withdraw()
    fTyp = [("","*")]
    iDir = os.path.abspath(os.path.dirname(__file__))
    tkinter.messagebox.showinfo('ディレクトリ選択','対象ディレクトリを選択してください！')

    dir = tkinter.filedialog.askdirectory(initialdir = iDir)


    # 処理ディレクトリパスの出力
    #tkinter.messagebox.showinfo('ディレクトリ選択',dir)

    return dir
