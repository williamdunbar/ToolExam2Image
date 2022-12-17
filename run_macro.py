from time import sleep
import sys
import win32com.client as win32
import argparse
import os


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
        print(sys.modules[wd.__module__].__file__)
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