from googletrans import Translator
import pandas as pd
from tkinter import *
from tkinter.filedialog import askopenfilename
import time
import threading

# Initialize the translator
file_path=''
total=0
stop_loop= False
translator = Translator()

# 파일 경로 받아오기
def getFile():
    global file_path
    file_path = askopenfilename(
        title='Select an Excel file',
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    file_path_value.config(text=file_path)

# 번역 메소드
def translateFun(df, row, index, lang_type, lang_source_type, column_name):
        global total
        global translator
        translator = Translator()
        translated = translator.translate(row[column_name], src=lang_source_type, dest=lang_type)
        df.loc[index, 'translation'] = translated.text
        translatedText =f'No {index+1}/{total}: {row[column_name]} --> {translated.text}'
        print(translatedText)
        progress_lab.config(text=translatedText)

# 변역 시작후 정지 기능
def stopLoop():
    global stop_loop
    stop_loop = True

# 파일 번역
def translateFile():
    # 총 행수
    global total
    # 파일경로
    global file_path 
    # 소스 언어코드
    lang_source_type = lang_source_input.get()
    # 번역할 언어코드
    lang_type = lang_type_input.get()
    # 번역할 엑셀 컬럼명
    column_name = column_input.get()

    global stop_loop
    stop_loop = False

    print(f'소스 언어코드:{lang_source_type}, 번역할 언어코드:{lang_type}, 번역할 엑셀 컬럼명{column_name}')

    if len(lang_source_type) == 0:
        lang_source_input.insert(0, '소스 언어코드를 입력하세요.')
        return
    
    if len(lang_type) == 0:
        lang_type_input.insert(0, '번역할 언어코드를 입력하세요.')
        return
    
    if len(column_name) == 0:
        column_input.insert(0, '번역할 엑셀 컬럼명을 입력하세요.')
        return
    
    if len(file_path) ==0:
        file_path_value.config(text="파일을 선택하세요.")
        return
    
    if file_path:
        df  =pd.read_excel(file_path, skiprows=2)
        print("Start translate--------------:")
        counter = 0 #번역 회수, 100회 번역시 잠시 sleep
        df['translation'] = 'default_value'
        df['translation'] = df['translation'].astype(str)
        total = len(df)
        finished_no = 0
        
        for index, row in df.iterrows():
            if stop_loop:
                break
            try:
                counter += 1
                translateFun(df, row, index, lang_type, lang_source_type, column_name)
                finished_no = index + 1
            except Exception as e:
                print(f"Error ----row index:{index}:{e}")
                print('sleep----10 SECONDS--for rest-------------')

                progress_lab.config(text=f"ERROR OCCURED: {e}")

                time.sleep(10)  # 10초 후 재시도
                translateFun(df, row, index, lang_type, lang_source_type, column_name )
                continue
            
            #100회 번역시 잠시 sleep
            if counter == 100:
                print(F'sleep---------{index+1} HAS BEEN FINISHED----------')
                time.sleep(20)  # 20초 후 재시도
                global translator
                translator = Translator()
                counter = 0
        if finished_no == total:
            print("Finished translation---------!")
            progress_lab.config(text="Translation Finished---------------!")
            df.to_excel('translated_file.xlsx', index=False)
        elif stop_loop:
            progress_lab.config(text=f"Stoped translation------------!")
        else:
            progress_lab.config(text=f"ERROR OCCURED----!")

    else:
        print("no file selected")


#-------------------------- widgets --------------------------
tk = Tk() 
tk.title('Translator Tool')

lang_source_lab = Label(tk,text='소스 언어코드 From:', width=16, anchor="w", relief="flat") 
lang_source_lab.grid(row=0, column=0, pady=(12, 4), padx=4)

lang_source_input =Entry(tk, width=20)
lang_source_input.insert(0, 'ko')
lang_source_input.grid(row=0, column=1,  columnspan=2, pady=(12, 4), padx=4)

lang_type_lab = Label(tk,text='번역할 언어코드 To:', width=16, anchor="w", relief="flat") 
lang_type_lab.grid(row=1, column=0, pady=(12, 4), padx=4)

lang_type_input =Entry(tk, width=20)
lang_type_input.grid(row=1, column=1,  columnspan=2, pady=(12, 4), padx=4)

column_lab = Label(tk,text='번역할 엑셀 컬럼명:', width=16, anchor="w", relief="flat") 
column_lab.grid(row=2, column=0, pady=(12, 4), padx=4)

column_input =Entry(tk, width=20)
column_input.insert(0, 'value')
column_input.grid(row=2, column=1,  columnspan=2, pady=(12, 4), padx=4)

file_lab = Label(tk,text='번역할 엑셀 파일:' , width=16, anchor="w") 
file_lab.grid(row=3, column=0, pady=4, padx=4)

upload_btn = Button(tk,text='파일 선택', bg='gray', fg='white',width=20, command=getFile)
upload_btn.grid(row=3, column=1, columnspan=2, pady=4, padx=4)

file_path_lab = Label(tk, text='파일 경로:',width=16, anchor="w") 
file_path_lab.grid(row=4, column=0,  pady=4, padx=4)                              

file_path_value = Label(tk, justify="left", wraplength=240) 
file_path_value.grid(row=5, column=0, columnspan=3, pady=4, padx=4)

devide_line = Label(tk, text='.................................................................................') 
devide_line.grid(row=6, column=0, columnspan=3, sticky="ew", padx=4, pady=4)

start_btn = Button(tk, bg='#04AA6D', fg='white', text="Start Translate")
start_btn.config(command=lambda:threading.Thread(target=translateFile).start())
start_btn.grid(row=7, column=0, columnspan=3, sticky="ew", padx=4, pady=4)

stop_btn = Button(tk, bg='#ff5757', fg='white', text="Stop Loop", command=stopLoop)
stop_btn.grid(row=8, column=0, columnspan=3, sticky="ew", padx=4, pady=4)

progress_lab = Label(tk, width=38, height=5, justify="left", wraplength=240, anchor="nw") 
progress_lab.config(text='번역정보 출력----------------:')
progress_lab.grid(row=9, column=0, columnspan=3, sticky="ew", padx=4, pady=10)

tk.mainloop()