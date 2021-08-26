#Extract_Hwp_Survey.py
#Made by sungcheol Ha
#E-mail : sungchol3@daum.net

#Import background
import tkinter as tk
from tkinter import messagebox, filedialog
import win32com.client as win32
import os
import re
from openpyxl import Workbook

#Open Hwp files and Scaning and get Data
class HwpFile():

    def __init__(self) -> None:
        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject") #Make Hwp controler
        self.hwp.XHwpWindows.Item(0).Visible = True #Visible HWP work
        self.hwp.RegisterModule("FilePathCheckDLL","SecurityModule") #Do work on background

    def Get_Survey_Result(self, file_list) -> list:
        data = self.Scan(file_list)
        self.hwp.Quit()
        xl_data = self.Organize(data)
        return xl_data
    
    def get_folder_link(self, file_list):
        result = list()
        for file in file_list:
            result.append(os.path.abspath(os.path.join(file)))

    def check_fileformat(self, filename):
        #Check file extension is hwp
        compiler = re.compile('\[(\d.*)\].*[.]hwp')
        return compiler.search(filename)
    
    def search(self, dirname) -> list:
        #Search hwp file and return its list
        filenames = os.listdir(dirname)
        result = []
        for filename in filenames:
            full_filename = os.path.join(dirname, filename)
            ext = os.path.splitext(full_filename)[-1]
            if self.check_fileformat(filename): result.append(filename)
        return result

    def Scan(self, file_list) -> list:
        data = []
        for file in file_list:
            self.hwp.Open(file)
            self.hwp.InitScan()
            result = list()
            while True:
                textdata = self.hwp.GetText()
                if textdata[0] == 1: break
                else:
                    #print(textdata[1])
                    result.append(textdata[1])
            self.hwp.ReleaseScan()
            print("filename : "+ os.path.split(file)[1])
            dict_data =self.organize_textlist_to_dict(result)
            data.append(self.make_excel_data(file, dict_data))
            print(dict_data)
        
        return data #Double List [[num,name,checknum]]
    
    def find_checknum(self, lst) -> int:
        #lst=[blank, blank, ..., 'v', blank]
        #return 1~5
        #find check point
        BLANK = ''
        for i, index in enumerate(lst):
            if index != BLANK: return (i+1)
        return None

    def remove_keystr(self, lst, key='\r\n') -> None:
        for i in range(len(lst)):
            lst[i] = lst[i].replace(key,'')

    def organize_textlist_to_dict(self,lst) -> dict:
        #clear_data
        result = []
        for i in range(1,18):
            try: first = lst.index("{}\r\n".format(i))
            except: print(lst)
            if i < 17: last = lst.index("{}\r\n".format(i+1))
            else: last = first + 7
            storage = lst[first:last]
            self.remove_keystr(storage)
            result.append(storage)
        
        #point_data
        dict_data = dict()
        for i, index in enumerate(result):
            j = self.find_checknum(index[-5:])
            dict_data[i+1] = j
        
        return dict_data

    def make_excel_data(self, file, dict_data) -> list:
        #make_excel_data(file, dic_data)
        compiler = re.compile("\[(\d{1,2})(.*)\].*[.]hwp")
        filename = os.path.split(file)[1]
        compiled = compiler.match(filename)
        student_name = compiled.group(2)
        student_num = int(compiled.group(1))
        survey_result = list(dict_data.values())
        survey_result.insert(0, student_num)
        survey_result.insert(1, student_name)
        return survey_result
    
    def Organize(self, dou_list) -> list:
        #make_blank_data(dou_list):
        #Return new ordered double list which make blank number student to fit format
        sorted_data = sorted(dou_list)
        result = list()
        before = 1
        for i, item in enumerate(sorted_data):
            num = item[0]
            if num != before:
                for j in range(before, num):
                    result.append([j])
            before = num+1
            result.append(item)
        return result

class ExcelFile():
    #Write Data to Excel File
    def __init__(self) -> None:
        self.xl = Workbook()
        self.sheet = self.xl.create_sheet('result')
    
    def save_to_excel(self, data, folder) -> None:
        for item in data:
            self.sheet.append(item)
        self.xl.save(os.path.join(folder, "result.xlsx"))

class Window():

    files = list() #Hwp Files List
    warning_message = "파일 이름이 정확하지 않은 파일이 존재합니다. 파일 이름은 [(연번)(이름)](제목).hwp 형태여야 합니다. 무시하고 계속하시겠습니까?"

    def __init__(self) -> None:
        #Set Window
        self.root = tk.Tk()
        self.root.title("KAIST 사이버영재교육 설문지 자동추합기")
        self.root.geometry("540x300+100+100")
        self.root.resizable(True,True)

        #Set Basic Format
        self.open = tk.Button(self.root, text="Open Files", command=self.OpenHwpFiles)
        self.open.pack(side="top")
        self.listbox = tk.Listbox(self.root, width=50)
        self.listbox.pack(side="top")
        self.deletebtn = tk.Button(self.root, text="Delete", command=self.delete_element)
        self.deletebtn.pack(side="top")
        self.okbtn = tk.Button(self.root, text="OK", command=self.Run)
        self.okbtn.pack(side="top")

        #Set Hwp and Excel
        #self.hwp = HwpFile()
        #self.excel = ExcelFile()

        #Run
        self.root.mainloop()

    def check_fileformat(self, filename):
        #Check file extension is hwp
        compiler = re.compile("\[(\d{1,2})(.*)\].*[.]hwp")
        return compiler.search(filename)

    def reset_listbox(self):
        self.files = list()
        self.listbox.delete(0,self.listbox.size())

    def show_files(self):
        for i, name in enumerate(self.files):
            self.listbox.insert(i, os.path.split(name)[1])
    
    def delete_element(self):
        i = self.listbox.curselection()[0]
        print(self.files.pop(i))
        self.listbox.delete(i,i)    

    def OpenHwpFiles(self):
        self.reset_listbox()
        nominee_list = filedialog.askopenfilenames(
            initialdir='path', 
            title='select files', 
            filetypes=(('hwp files (*.hwp, *.hwpx)', ('*.hwp','*.hwpx')), 
                ('all files', '*.*')))
        
        for name in nominee_list:
            if self.check_fileformat(name) != None:
                self.files.append(name)
            else:
                messagebox.showwarning("Unmatch file name",self.warning_message)
        self.show_files()
    
    def Run(self):
        #Set Hwp and Excel editor
        self.hwp = HwpFile()
        self.excel = ExcelFile()

        #Scan HWP and get Excel Data
        data = self.hwp.Get_Survey_Result(self.files)

        #Save at Excel
        folder = filedialog.askdirectory(title="결과를 저장할 폴더를 선택하세요")
        self.excel.save_to_excel(data, folder)
        messagebox.showinfo("Complete Save", os.path.join(folder,"result.xlsx")+"에 성공적으로 저장되었습니다")


if __name__ == "__main__":
    window = Window()