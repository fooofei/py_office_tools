# coding=utf-8



import os
import sys
from datetime import datetime


class ExcelSecurityRegWriter():
    '''
    要调用 COM 在 xls 中写入 vba 需要 开启 注册表权限
    '''

    def __init__(self):
        import win32con
        self.key = win32con.HKEY_CURRENT_USER
        self.regpath = r"Software\Microsoft\Office\15.0\Excel\Security"
        self.valuename = "AccessVBOM"
        self.valuetype = win32con.REG_DWORD

    def __del__(self):
        self.close_access()

    def _modify_access(self, key, keypath, valuename, valuetype, value):
        import win32con
        import win32api
        try:
            keyhandle = win32api.RegConnectRegistry(None, key)
            subkeyhandle = win32api.RegOpenKeyEx(keyhandle, keypath, 0, win32con.KEY_READ)
            curvalue, type = win32api.RegQueryValueEx(subkeyhandle, valuename)
            if curvalue != value:
                win32api.RegCloseKey(subkeyhandle)
                subkeyhandle = win32api.RegOpenKeyEx(keyhandle, keypath, 0, win32con.KEY_SET_VALUE)
                win32api.RegSetValueEx(subkeyhandle, valuename, 0, valuetype, value)
            win32api.RegCloseKey(subkeyhandle)
            win32api.RegCloseKey(keyhandle)
            return True
        except:
            return False

    def open_access(self):
        return self._modify_access(self.key, self.regpath, self.valuename, self.valuetype, 1)

    def close_access(self):
        self._modify_access(self.key, self.regpath, self.valuename, self.valuetype, 0)


def echo_macro_content_to_excel(fullpath_excel, macro_content):
    import win32com.client as win32
    regobj = ExcelSecurityRegWriter()
    if regobj.open_access():
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(Filename=fullpath_excel)
        xlmodule = workbook.VBProject.VBComponents.Add(1)
        xlmodule.CodeModule.AddFromString(macro_content)
        excel.Workbooks(1).Close(SaveChanges=1)
        excel.Application.Quit()
        del excel
    else:
        raise ValueError(u"Error : Reg Failed, cannot write excel")


def create_new_excel(fullpath_dest):
    import xlwt

    style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
                         num_format_str='#,##0.00')
    style1 = xlwt.easyxf(num_format_str='D-MMM-YY')
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')
    ws.write(0, 0, 1234.56, style0)
    ws.write(1, 0, datetime.now(), style1)
    try:
        wb.save(fullpath_dest)
        return True
    except:
        return False


def write_macro_content_to_random_file(fullpath_macro):
    import time

    with open(fullpath_macro, "rb") as fp:
        content = fp.read()
        if content:
            name = u"{}_generate.xls".format(time.strftime(u'%Y_%m_%d_%H_%M_%S', time.localtime()))
            path_new = fullpath_macro + name
            if os.path.exists(path_new):
                os.remove(path_new)
            if create_new_excel(path_new):
                echo_macro_content_to_excel(path_new, content)
                return path_new
    return None


def entry():
    import io_in_out
    fs = io_in_out.io_iter_files_from_arg(sys.argv[1:])
    for e in fs:
        write_macro_content_to_random_file(e)
    raw_input('enter...')


if __name__ == "__main__":
    entry()
