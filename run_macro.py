from time import sleep
import win32com.client as win32
import argparse
import os





# def init_parser():
#     parser = argparse.ArgumentParser()
#     parser.add_argument('macro_name', help='name of the macro to run in Word. Make sure the macro exists.')
#     parser.add_argument('-r', '--recurse', action='store_true', help='recursively search through subfolders when looking for docx files. (default: false)')
#     parser.add_argument('-s', '--suffix', help='customize the suffix for processed files. (default: _done)')
#     parser.add_argument('doc_path', nargs=argparse.REMAINDER, help='path/s to .docx files or a folders that contains .docx files. (default: current folder)')
#     parser.epilog = "© 2021 - Paul McClintock - https://gist.github.com/plauk/1d26b3c6434a6b0fc44a9b213bf92d77"
#     parser.usage = "run_macro.py [-h] [-r] [-s SUFFIX] macro_name *doc_path (*add as many as you want)"
#     args = parser.parse_args()
#     return args


def run_macros(wd, doc_path, macro_name, save_as_suffix, recurse):
    dp = doc_path
    dp = os.path.abspath(dp) # convert to absolute path
    # if path is a folder, grab ALL docx files from folder
    # otherwise, just run this file

    if os.path.isdir(dp):
        # recursively (or not) call this function for all docx files in this folder
        sub_items = [os.path.join(dp, f) for f in os.listdir(dp) if f.endswith("docx") or os.path.isdir(os.path.join(dp,f))] if recurse else [os.path.join(dp, f) for f in os.listdir(dp) if f.endswith("docx")]
        run_macros(wd, sub_items, macro_name, save_as_suffix, recurse)  # recurse
    else:
        # open doc
        doc = wd.Documents.Open(dp)
        # pth = doc.Path + wd.PathSeparator
        pth = os.getcwd()
        nm = doc.Name.split(".")[0]
        ext = "." + doc.Name.split(".")[1]
        print(f"Opened {nm}", f" from {pth}")

        # run macro
        print("Attempting to run macro: ", macro_name)
        wd.Application.Run(macro_name)

        # save as, close and clean up
        print("saving and closing doc", nm)
        # doc.SaveAs(pth + nm + save_as_suffix + ext)
        print(pth+"\Macro.docx")
        doc.SaveAs(pth+"\Macro.docx")
        doc.Close()
        doc = None

def RunMacro(doc_path):
    try:
        recurse = False
        save_as_suffix = "_done"
        macro_name = "W2T_Tool"
        if len(doc_path) < 1:
            doc_path = ["."]

        wd = win32.gencache.EnsureDispatch("Word.Application")
        wd.Visible = False
        print(str(doc_path))
        os.startfile(str(doc_path))
        sleep(5)
        run_macros(wd, doc_path, macro_name, save_as_suffix, recurse)

        # clean up
        print("Closing Word")
        wd.Quit()
        wd = None
    except Exception as err:
        return str(err), "Đã có lỗi khi thực hiện macro!"
    return "", ""