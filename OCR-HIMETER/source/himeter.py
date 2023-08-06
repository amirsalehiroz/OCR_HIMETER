from customtkinter import *
from tkinter import Menu
from tkinter import *
from PIL import Image
from tkinter.filedialog import askdirectory
from tkinter import StringVar
from tkinter import Checkbutton
import csv
import pandas as pd
import os
import pytesseract
import aspose.pdf as ap

class _Main_:
    def __init__(self):
        # --------------------------------------
        def get_file():
            files = askdirectory(initialdir="/", title="Select Files Directory")
            self.get_file_path.insert(0, files)
            with open('get_file.csv', 'w')as file:
                csv_data = csv.writer(file)
                csv_data.writerow(['get_file_path'])
                data = [str(files)]
                csv_data.writerows([data])
                
        def output_file():
            files = askdirectory(initialdir="/", title="Select Output Directory")
            self.get_output_path.insert(0, files)
            with open('get_output.csv', 'w')as file:
                csv_data = csv.writer(file)
                csv_data.writerow(['get_output_path'])
                data = [str(files)]
                csv_data.writerows([data])

        def run():

            file_path = pd.read_csv('get_file.csv')['get_file_path'].iloc[0]
            output_path = pd.read_csv('get_output.csv')['get_output_path'].iloc[0]
            file_save_format = self.format_option_menu.get()
            save_file_parametrs = self.save_file_parametrs.get()
            language = self.language.get()
            os.chdir(file_path)
            listdir = os.listdir()
            direct_path = os.getcwd()
            if language == 'Persian':
                lang = 'fas'
            if language == 'English':
                lang = 'eng'
            if language == 'Arabic':
                lang = 'ara'
            if language == 'China':
                lang = 'cha'

            if file_save_format == 'DOC':
                format = 'doc'
            if file_save_format == 'PDF':
                format = 'pdf'
            if file_save_format == 'MD':
                format = 'md'
            if file_save_format == 'TXT':
                format = 'txt'
            if file_save_format == 'HTML':
                format = 'html'
            
          
            if save_file_parametrs == 'Save the entire text in one file':
                if  file_save_format=='PDF':
                        for i in listdir:
                            os.chdir(file_path)
                            listdir = os.listdir()
                            direct_path = os.getcwd()
                            img = Image.open(f'{direct_path}\{i}')
                            pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
                            result = pytesseract.image_to_string(img, lang=(f'{lang}'))
                            os.chdir(output_path)
                            document = ap.Document()
                            page = document.pages.add()

                            if lang == 'fas':
                                position = ap.HorizontalAlignment.RIGHT
                            elif lang == 'eng':
                                position = ap.HorizontalAlignment.LEFT

                            text_fragment = ap.text.TextFragment(result)
                            text_fragment.horizontal_alignment = position
                            text_fragment.text_state.font = ap.text.FontRepository.find_font("Arial")
                            text_fragment.text_state.font_size = 24
                            page.paragraphs.add(text_fragment)

                            document.save(f'{i}.pdf')
                            self.show_box.insert(1.0, result)
                else:
                            for i in listdir:
                                os.chdir(file_path)
                                listdir = os.listdir()
                                direct_path = os.getcwd()
                                img = Image.open(f'{direct_path}\{i}')
                                pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
                                result = pytesseract.image_to_string(img, lang=(f'{lang}'))
                                os.chdir(output_path)
                                with open(f'{i}.{format}', 'w', encoding='utf-8')as file:
                                    file.write(result)
                                self.show_box.insert(1.0, result)

            if save_file_parametrs == 'Saving texts separately':
                if  file_save_format=='PDF':
                        for i in listdir:
                            os.chdir(file_path)
                            listdir = os.listdir()
                            direct_path = os.getcwd()
                            img = Image.open(f'{direct_path}\{i}')
                            pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
                            result = pytesseract.image_to_string(img, lang=(f'{lang}'))
                            os.chdir(output_path)
                            document = ap.Document()
                            page = document.pages.add()

                            if lang == 'fas':
                                position = ap.HorizontalAlignment.RIGHT
                            elif lang == 'eng':
                                position = ap.HorizontalAlignment.LEFT

                            text_fragment = ap.text.TextFragment(result)
                            text_fragment.horizontal_alignment = position
                            text_fragment.text_state.font = ap.text.FontRepository.find_font("Arial")
                            text_fragment.text_state.font_size = 24
                            page.paragraphs.add(text_fragment)

                            document.save(f'{i}.pdf')
                            self.show_box.insert(1.0, result)
                else:
                            for i in listdir:
                                os.chdir(file_path)
                                listdir = os.listdir()
                                direct_path = os.getcwd()
                                img = Image.open(f'{direct_path}\{i}')
                                pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'
                                result = pytesseract.image_to_string(img, lang=(f'{lang}'))
                                os.chdir(output_path)
                                with open(f'{i}.{format}', 'w', encoding='utf-8')as file:
                                    file.write(result)
                                self.show_box.insert(1.0, result)


        def clear_box():
            self.get_file_path.delete(0 , END)
            self.get_output_path.delete(0 ,END)
            self.show_box.delete(1.0 , END)
            self.log_box.delete(1.0 ,END)
        def text_viewver():
            self.root.geometry('900x550')
        def hide():
            self.root.geometry('460x550')

        #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        # --------------------------------------
        self.root = CTk()
        self.root.geometry('460x550')
        self.root.resizable(False , False)
        self.root.iconbitmap('../Image/icon.ico')
        self.root.title('OCR HIMETER')
        set_appearance_mode('light')
        self.menubar = Menu(self.root)
        file_menu = Menu(self.menubar, tearoff=0)
        file_menu.add_command(label='New')
        file_menu.add_command(label='Export')
        file_menu.add_separator()
        file_menu.add_command(label='Exit', command=self.root.quit)
        self.menubar.add_cascade(label="File", menu=file_menu)
        #------------------------------------------------------
        help_menu = Menu(self.menubar , tearoff=0)
        help_menu.add_command(label='Tutorial')
        help_menu.add_command(label='About')
        self.menubar.add_cascade(label='Help', menu=help_menu)
        #-------------------------------------------------------
        self.root.configure(menu=self.menubar)
        self.main_frame = CTkFrame(master=self.root)
        self.main_frame.pack(fill=BOTH, expand=True)

        self.ocr_himeter = CTkLabel(master=self.main_frame, text='OCR HIMETER', font=('Arial', 20), text_color='black', )
        self.ocr_himeter.place(x=175, y=20)
        self.find_path = CTkImage(dark_image=Image.open("../Image/left_arrow.png"), size=(20, 20))

        self.find_path_label = CTkLabel(master=self.main_frame , text='Path Folder', font=('Arial', 15), text_color='black')
        self.find_path_label.place(x=50, y=80)

        self.get_file_path = CTkEntry(master=self.main_frame ,placeholder_text='Enter Path Address',width=200, font=('Arial', 15), text_color='black')
        self.get_file_path.place(x=145, y=80)

        self.get_file_path_button = CTkButton(master=self.main_frame,image=self.find_path, text='', width=20, fg_color='green', command=get_file)
        self.get_file_path_button.place(x=350, y=80)

        self.format_label = CTkLabel(master=self.main_frame, text='Format', font=('Arial', 15), text_color='black')
        self.format_label.place(x=50, y=140)

        self.format_option_menu = CTkOptionMenu(master=self.root,width=140,text_color ='black',fg_color='white',bg_color='black' ,dynamic_resizing=False,
                                                        values=["Select Format" ,"DOC", "PDF", "MD", "TXT", "HTML"])
        self.format_option_menu.place(x=120, y=140)

        self.save_in_onefile = CTkLabel(master = self.main_frame, text='File save type' ,text_color='black', font=('Arial',15))
        self.save_in_onefile.place(x=50, y=200)

        self.save_file_parametrs = CTkOptionMenu(master=self.main_frame,width=205,text_color ='black',fg_color='white',bg_color='black' ,dynamic_resizing=False,
                                                        values=["Select Format" ,"Save the entire text in one file", "Saving texts separately"])
        self.save_file_parametrs.place(x=155, y=200)

        self.output_path_label = CTkLabel(master=self.main_frame, text='Output Folder', text_color='black', font=('Arial',15))
        self.output_path_label.place(x=50, y=260)

        self.get_output_path = CTkEntry(master=self.main_frame ,placeholder_text='Enter Output Address',width=200, font=('Arial', 15), text_color='black')
        self.get_output_path.place(x=158, y=260)

        self.get_output_path_button = CTkButton(master=self.main_frame,image=self.find_path, text='', width=20, fg_color='green', command=output_file)
        self.get_output_path_button.place(x=364, y=260)

        self.create_button = CTkButton(master=self.main_frame, text='Create', text_color='black',width=70,fg_color='green', command=run)
        self.create_button.place(x=180, y=370)

        self.log_box = CTkTextbox(master=self.main_frame, width=460, font=('Arial',20), border_width=2, border_color='green', padx=10)
        self.log_box.place(x=0, y=410)
        self.log_box.configure(text_color='red')
        self.log_box.insert(1.0, 'This app for convert png/jpg file to text file (PDF/MD/TXT/HTML/WORD)\nTelegram:@abolfazl_salehigg\nGitHub:https://github.com/amirsalehiroz\nPowerd by Abolfaz Salehi')
  

        self.show_box = CTkTextbox(master=self.main_frame, width=440, font=('Arial', 15), height=600, border_width=2, border_color='black', pady=7, padx=6)
        self.show_box.place(x=460, y=0)

        self.clear_button = CTkButton(master=self.main_frame, text='Clear', text_color='black',width=70,fg_color='green', command=clear_box)
        self.clear_button.place(x=100, y=370)

        self.textviewer_button = CTkButton(master=self.main_frame, text='Text Viewer', text_color='black',width=70,fg_color='green', command=text_viewver)
        self.textviewer_button.place(x=261, y=370)

        self.hide_button = CTkButton(master=self.main_frame, text='Hide', text_color='black',width=70,fg_color='green', command=hide)
        self.hide_button.place(x=465, y=521)
        
        self.language_label = CTkLabel(master=self.main_frame, text='Language', text_color='black', font=('Arial',15))
        self.language_label.place(x=50, y=320)

        self.language = CTkOptionMenu(master=self.main_frame,width=160,text_color ='black',fg_color='white',bg_color='black' ,dynamic_resizing=False,
                                                        values=["Select Language","Persian" ,"English", "Arabic","China"])
        self.language.place(x=150, y=320)

        #+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



        self.root.mainloop()

if __name__ == '__main__':
        _Main_()
