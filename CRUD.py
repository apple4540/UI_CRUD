import pyodbc
from tkinter import ttk,filedialog, messagebox
import tkinter as tk
import customtkinter
import tksheet
from CTkTable import *
import pandas as pd
from pandastable import Table
import openpyxl

#customtkinter.set_appearance_mode("system")#默認主題外觀
#customtkinter.set_default_color_theme("blue")#主題顏色

app=customtkinter.CTk()#建立視窗
app.geometry("900x400")#視窗大小設定
app.title('CRUD')#視窗標題
frame =customtkinter.CTkFrame(app, fg_color='#FFEEDD')#變更背景顏色
frame.pack(fill='both', expand=True)#填滿整個視窗

#連接資料庫
connection=pyodbc.connect('Driver={SQL server};\
                                    Server=WANCENLI-NB\SQLEXPRESS;\
                                    Database=test;\
                                    UID=sa;\
                                    PWD=abc123'
                                    #Trusted_connection=True'
                                    )
connection.autocommit=True# SQL 操作在執行後會自動被提交，因此不需要呼叫 connection.commit()


####設置各個欄位的輸入框####
# entry_table_name=customtkinter.CTkEntry(app,placeholder_text="Table Name")
# entry_table_name.place(relx=0.2,rely=0.1)

# region###新增/刪除資料輸入框與標籤####
Batch_Number_var = tk.StringVar()#設置一個【電芯批號】變數
Batch_Number_var.set("電芯批號：")
Batch_Number_label = customtkinter.CTkLabel(master=app,textvariable=Batch_Number_var,
                                            fg_color='#FFEEDD',#更改背景顏色
                                            font=('Microsoft JhengHei', 13,'bold'),#更改字體與大小
                                            text_color='#842B00'# 设置文字颜色
                                            )
Batch_Number_label.place(x=10,y=30)
Batch_Number_var_entry=tk.StringVar()#設置一個【電芯批號輸入】變數
entry_Batch_Number=customtkinter.CTkEntry(app,textvariable=Batch_Number_var_entry)#用於點選表格後自動代入至電芯批號輸入框
entry_Batch_Number.place(x=100,y=30)

Owner_var= tk.StringVar()
Owner_var.set("負責人：")
Owner_label = customtkinter.CTkLabel(master=app,textvariable=Owner_var,
                                    fg_color='#FFEEDD',#更改背景顏色
                                    font=('Microsoft JhengHei', 13,'bold'),#更改字體與大小
                                    text_color='#842B00'# 设置文字颜色
                                    )
Owner_label.place(x=10,y=70)
Owner_var_entry=tk.StringVar()
entry_Owner=customtkinter.CTkEntry(app,textvariable=Owner_var_entry)
entry_Owner.place(x=100,y=70)

Project_Number_var= tk.StringVar()
Project_Number_var.set("專案編號：")
Project_Number_label = customtkinter.CTkLabel(master=app,textvariable=Project_Number_var,
                                     fg_color='#FFEEDD',#更改背景顏色
                                    font=('Microsoft JhengHei', 13,'bold'),#更改字體與大小
                                    text_color='#842B00'# 设置文字颜色
                                    )
Project_Number_label.place(x=10,y=110)
Project_Number_var_entry=tk.StringVar()
entry_Project_Number=customtkinter.CTkEntry(app,textvariable=Project_Number_var_entry)
entry_Project_Number.place(x=100,y=110)



###刪除資料輸入框###
d_Batch_Number_var = tk.StringVar()#設置一個要刪除的【電芯批號】變數
d_Batch_Number_var.set("電芯批號：")
entry_d_Batch_Number=customtkinter.CTkEntry(app,placeholder_text="Delete Batch Number")
entry_d_Batch_Number.place(x=250,y=30)
# endregion

class Select:#選取物件定義
    def on_tree_select(event):#定義選取表格上的資料並代入至輸入框
    # 取得選擇的項目
        selected_item = tree.selection()
        if selected_item:
            #取得選擇項目中的值 
            item = tree.item(selected_item)
            values = item['values']
            if values:
                #print(f"Selected item: {values}")
                # 選取的值填入 Entry 
                Batch_Number_var_entry.set(values[0])
                Owner_var_entry.set(values[1])
                Project_Number_var_entry.set(values[2])

    def on_combobox_select(event):#下拉選單選擇後開啟新視窗
        selected_option = combobox.get()
        if selected_option:
            # 建立新的視窗
            new_window = tk.Toplevel(app)
            new_window.title(f"{selected_option}")
            new_window.geometry("300x200")
            new_window.configure(bg="#FFEEDD")  # 设置背景颜色
            test_label = tk.Label(new_window, text=f"{selected_option}",
                                   font=("Helvetica", 16),
                                   #fg_color='#FFEEDD',#更改背景顏色
                                   #text_color='black'
                                   )
            test_label.pack(pady=20)

class up_down_load:#定義上傳下載
    def display_data(df):
        for widget in app.winfo_children():
            widget.destroy()

        cols = list(df.columns)
        tree = ttk.Treeview(app, columns=cols, show="headings")
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, minwidth=0, width=100, stretch=False)
        
        for index, row in df.iterrows():
            tree.insert("", "end", values=list(row))
    def upload_file():
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)
                up_down_load.display_data(df)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file\n{e}")

class CRUD:
    def insert():#定義輸入新增新資料
        try:
            cursor=connection.cursor()#目前執行位置
            cursor.execute(#f"INSERT INTO {entry_table_name.get()} VALUES"+
                            "INSERT INTO Coin_cell VALUES"+
                            f"('{entry_Batch_Number.get()}',"+ #文字用''標示，f轉字串
                            f"'{entry_Owner.get()}',"+
                            f"'{entry_Project_Number.get()}')"
                                )
            info_label.configure(text="INSERT COMPLETED!")#輸入成功顯示訊息
            cursor.close()
        except pyodbc.Error as ex:
            print('Connection failed',ex)
            info_label.configure(text="INSERT FAILED!")

    def delete():#定義刪除資料
        try:
            cursor=connection.cursor()#目前執行位置
            cursor.execute("DELETE FROM Coin_cell WHERE [Batch Number]="+
                           f"('{entry_d_Batch_Number.get()}')"
                           )
            info_label.configure(text="DELETE COMPLETED!")#輸入成功顯示訊息
            cursor.close()
        except pyodbc.Error as ex:
            print('Connection failed',ex)
            info_label.configure(text="DELETE FAILED!")

    def update():#定義更新資料，以電芯批號作為更新依據
        try:
            cursor=connection.cursor()#目前執行位置
            Batch_Number=entry_Batch_Number.get()
            Owner=entry_Owner.get()
            Project_Number=entry_Project_Number.get()
            cursor.execute("UPDATE Coin_cell "+
                           "SET Owner = '%s', [Project Number]= '%s' WHERE [Batch Number] = '%s' " %(Owner,Project_Number,Batch_Number))
            info_label.configure(text="UPDATE COMPLETED!")#輸入成功顯示訊息
            cursor.close()
        except pyodbc.Error as ex:
            print('Connection failed',ex)
            info_label.configure(text="UPDATE FAILED!")

    def read():#定義讀取資料用以重新整理表格
        cursor=connection.cursor()#目前執行位置
        for row in tree.get_children():
            tree.delete(row)
        cursor.execute('SELECT * FROM Coin_cell')
        for row in cursor.fetchall():
            cleaned_row = [str(item).replace('(', '').replace(')', '').replace("'", "") for item in row]
            tree.insert('', tk.END, values=cleaned_row)
        cursor.close()
    
    def refresh_table():#定義重整表格
        CRUD.read()
    
    #備用    
    #def read():#讀取資料庫資料#回傳dataframe
        # try:
        #     cursor.execute("SELECT * FROM Coin_cell")#執行讀取指令
        #     info_label.configure(text="READ COMPLETED!")#讀取完成顯示訊息
        #     field_name=[i[0] for i in cursor.description]
        #     result =cursor.fetchall()

        #     data = pd.DataFrame()
        #     for row in result:
        #         data=pd.DataFrame(columns=field_name,index=list(range(1,10)))
        # except pyodbc.Error as ex:
        #     print('Connection failed',ex)
        #     info_label.configure(text="READ FAILED!")
        # try:
        #     cursor.execute('SELECT * FROM Coin_cell')
        #     field_name=[i[0] for i in cursor.description]#顯示column名稱
        #     #result =cursor.fetchall()
        #     result=pd.read_sql('SELECT * FROM Coin_cell',connection)#讀取SQL資料並轉成dataframe
        #     return field_name         
        # except pyodbc.Error as ex:
        #     print('Connection failed',ex)
        #     info_label.configure(text="READ FAILED!")

    


# region#建立按鈕
# #新增資料按鈕
insert_button=customtkinter.CTkButton(app,text="新增資料",
                                      command=CRUD.insert,
                                      fg_color="#F75000",
                                      text_color='white',
                                      font=('Microsoft JhengHei', 13,'bold')#更改字體與大小
                                      )
insert_button.place(x=100,y=150)#按鈕座標位置
#建立刪除按鈕
delete_button=customtkinter.CTkButton(app,text="刪除資料",
                                      command=CRUD.delete,
                                      fg_color="#FF8F59",
                                      text_color='white',
                                      font=('Microsoft JhengHei', 13,'bold')#更改字體與大小
                                      )
delete_button.place(x=250,y=150)
#建立更新按鈕
update_button=customtkinter.CTkButton(app,text="更新資料",
                                      command=CRUD.update,
                                      fg_color="#F75000",
                                      text_color='white',
                                      font=('Microsoft JhengHei', 13,'bold')#更改字體與大小
                                      )
update_button.place(x=400,y=150)
#建立重新整理表格按鈕
refresh_button=customtkinter.CTkButton(app,text="重新整理",
                                      command=CRUD.refresh_table,
                                      fg_color="#FF8F59",
                                      text_color='white',
                                      font=('Microsoft JhengHei', 13,'bold')#更改字體與大小
                                      )
refresh_button.place(x=550,y=150)
#建立上傳excel按鈕
upload_button = customtkinter.CTkButton(app,text="上傳",
                                      command=up_down_load.upload_file,
                                      fg_color="#F75000",
                                      text_color='white',
                                      font=('Microsoft JhengHei', 13,'bold')#更改字體與大小
                                      )
upload_button.place(x=700,y=150)
# endregion

info_label=customtkinter.CTkLabel(app,text="")#顯示目前狀態，EX:輸入成功
info_label.place(x=100,y=180)


# region#建立table
columns =('Batch Number', 'Owner', 'Project Number')#取得dataframe中的欄位名稱
tree = ttk.Treeview(app, columns=columns, show='headings')
tree.config(style="Custom.Treeview")
for col in columns:
     tree.heading(col, text=col)
     tree.column(col, width=100)

# 建立垂直滾動條
vsb = ttk.Scrollbar(app, orient='vertical', command=tree.yview)
vsb.place(x=10+1100,#Table座標x=10+寬度=1100
          y=300,# y=300座標
          height=200 #長度=200
          )
tree.configure(yscrollcommand=vsb.set)
# 建立水平滾動條
hsb = ttk.Scrollbar(app, orient='horizontal', command=tree.xview)
hsb.place(x=10,# x=10座標
          y=300+200,#y=300+高度=200
          width=1100#寬度=1100
          )
tree.configure(xscrollcommand=hsb.set)

#美化table
style_table = ttk.Style()
style_table.theme_use('clam')
style_table.configure("Custom.Treeview.Heading", font=('Helvetica',15, 'bold'), background="#842B00", foreground="white")
style_table.configure("Custom.Treeview", font=('Helvetica',13), rowheight=25)#表格內數值字體調整


tree.place(x=10, y=300, width=1100, height=200)#表格大小調整

tree.bind("<Double-1>",Select.on_tree_select)# 綁定 Treeview 的選擇事件

# endregion

# region# 建立下拉選單
style_combobox = ttk.Style()#建立樣式
style_combobox.theme_use('clam')#下拉選單樣式調整
style_combobox.configure("TCombobox",#默認字體無法調整大小(待確認有無其他方法)
                 font=('Microsoft JhengHei', 50),
                 foreground="#842B00",
                 padding=5#內邊距大小
                   )
combobox = ttk.Combobox(app, values=["表單 1", "表單 2", "表單 3"],style="TCombobox")
combobox.set("表單選擇")  # 默認顯示文字
combobox.place(x=100, y=15, anchor='center')  # combobox擺放位置

combobox.option_add('*TCombobox*Listbox.font', ('Microsoft JhengHei', 12))#下拉選單文本字體樣式調整
combobox.option_add('*TCombobox*Listbox.foreground', '#842B00')

combobox.bind("<<ComboboxSelected>>", Select.on_combobox_select)

# endregion



CRUD.read()#讀取表格資料

# table = CTkTable(master=app, row=5, column=5, values=CRUD.read())
# table.pack(expand=True, fill="both", padx=20, pady=20)
# def tabview():
#     frame=tk.Frame(app)
#     frame.pack(fill="both",expand=True)
#     table=Table(frame)
#     table.model.df=CRUD.read()

app.mainloop()#運行應用程序的主循環
  
