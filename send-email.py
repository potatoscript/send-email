part A   import os
import socket
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from python.codepython123 import Template as template
from datetime import datetime

class Views:
    def __init__(self, root):
        self.root = root
        self.window_width = 1200
        self.window_height = 600
        try:
            self.template = template()
            self.template.set_window_position(root,root,self.window_width,self.window_height)
            
            self.image_references = {}  # Dictionary to store image references
            self.image_references['reload_icon'] = self.template.get_icon('reload')
            self.image_references['delete_icon'] = self.template.get_icon('delete')
            self.image_references['master_icon'] = self.template.get_icon('master')
            self.image_references['csv_icon'] = self.template.get_icon('csv')
            self.image_references['save_icon'] = self.template.get_icon('save')
            self.image_references['register_icon'] = self.template.get_icon('register')
            self.image_references['add_icon'] = self.template.get_icon('add')
            self.image_references['send_icon'] = self.template.get_icon('send')
            self.image_references['preview_icon'] = self.template.get_icon('preview')
            self.image_references['complete_icon'] = self.template.get_icon('complete')

            self.csv_path="./"
            
            self.localization = {
                    "title": f"社内メール ( VERSION: Ver 2.3.0 ) ",
                    "boxno_label":"件名検索:",
                    "reload_button":"一覧更新",
                    "refresh_button":"更新",
                    "delete_button":"削除",
                    "add_button":"新規",
                    "meeting_minutes_button":"議事録",
                    "daily_report_button":"日報",
                    "businesstrip_button":"出張連絡",
                    "send_button":"送信",
                    "preview_button":"プレビュー",
                    "temporary_save_button":"一時保存",
                    "logout_button":"ログアウト",
                    "register_button": "復元",
                    "exported_to_csv": "データがCSVに正常にエクスポートされました。",
                    "login_window_title":"ユーザーログイン PC:",
                    "user_id":"ユーザーＩＤ:",
                    "user_level":"ユーザー権限:",
                    "login_button":"ログイン",
                    "logout_message":"がログアウトしました",
                    "login_message":"がログインしました",
                    "error":"エラー",
                    "error_password":"誤ったパスワードです",
                    "info":"お知らせ",
                    "no_userid":"ユーザーＩＤを入力してください。",
                    "no_boxno":"BoxNoを入力してください。",
                    "master_window_title":"ユーザーＩＤ登録",
                    "input_window_title":"メール項目入力：",
                    "save_button":"登録",
                    "password":"パスワード",
                    "master_password_input":"登録画面表示するにはパスワードを入力してください：",
                    "wrong_password":"誤ったパスワードです。",
                    "master_update_confirm":"ユーザーＩＤを更新してもよろしいですか？",
                    "master_add_confirm":"Global IDを追加してもよろしいですか？",
                    "confirm":"確認",
                    "ring_id":"リングID",
                    "type_code":"ﾃﾞﾊﾞｲｽﾀｲﾌﾟ",
                    "ring_good_qty":"ﾘﾝｸﾞ毎良品数",
                    "ring_bad_qty":"ﾘﾝｸﾞ毎不良品数", 
                    "delivery_date":"出荷予約日", 
                    "semi_make_date":"SEMIﾌｧｲﾙ作成日", 
                    "map_send_date":"Mapﾃﾞｰﾀ転送日",
                    "no_data":"データがありません",
                    "none_integer_error":"BoxNoは数字である必要があります",
                    "backup_complete":"バックアップ完了",
                    "register_complete":"復元完了",
                    "table":"ﾃｰﾌﾞﾙ",
                    "data_delete_confirm":"対象データを削除してもよろしいですか？",
                    "all_data_delete_confirm":"全てのデータを削除してもよろしいですか？",
                    "delete_confirm":"を削除してもよろしいですか？",
                    "password_input_message":"続行するにはパスワードを入力してください：",
                    "created":"追加した",
                    "updated":"更新した",
                    "deleted":"削除した",
                    "all_deleted":"全てデータを削除した",
                    "wrong_password_message":"誤ったパスワードです。削除はキャンセルされました。",
                    "canceled":"キャンセルされました",
                    "moved_to_cancel_folder":"のworkfolderは「出荷取り消し」に移動されました.",
                    "moved_back_send_folder":"のworkfolderは「出荷済み」に移動されました.",
                    "delivery_canceled":"出荷取り消した",
                    "master_chk":"登録",
                    "write_chk":"削除",
                    "read_chk":"閲覧",
                    "no_read_chk":"「閲覧」 チェックボックスをチャックしてください。",
                    "register_button":"一般者ログイン",
                    "register_window_title":"ユーザー情報登録",
                    "processing_backup":"データをバックアップ中．．．",
                    "processing_delete":"データを削除中．．．",
                    "processing_restore":"データを復元中．．．",
                    "recipient":"送信先",
                    "send_date":"日付", 
                    "subject":"件名", 
                    "message":"本文",
                    "message_work_content":"業務内容",
                    "message_work_today":"【本日】",
                    "message_work_nextday":"【翌日】",
                    "message_comment":"【所感】",
                    "register_button":"ユーザー登録",
                    "user_notfound":"未登録のユーザーです。先にユーザー登録して下さい。",
                    "user_email":"送信元メール",
                    "user_name":"送信元",
                    "user_department":"送信元部署",
                    "send":"送信",
                    "data_send_confirm":"対象メールを送信してもよろしいですか？",
                    "data_resend_confirm":"対象メールは送信しました、再送信してもよろしいですか？",
                    "data_send":"対象メールを送信しました",
                    "send_completed":"送信済",
                    "data_send_complete_confirm":"対象メールを送信済に処理してもよろしいですか？",
                    "data_send_complete_updated":"対象メールを送信済に処理しました",
                    "check_your_email":"自分のメールをご確認ください。",
                    "header_meeting_minutes":"ヘッダーその他",
                    "header_daily_report":"ヘッダー日報",
                    "saved":"保存しました"
                }

            self.title = self.localization.get("title")
            self.root.title(self.title)

            self.master_new = True  # True = insert data,  False = update data

            self.data_id = 0
            self.sender = ''
            self.send = 0
            self.type = 'daily_report or meeting_minutes'

            self.globalid = os.getlogin().upper()
            self.pc_name = socket.gethostname()
            
            self.create_login_window()
            
        except Exception as error:
            messagebox.showerror("Error", f"An error occurred: {error} - Connection to the server has not been established yet.")


    def create_main_section(self):
        input_frame = ttk.Frame(self.root)
        input_frame.pack(pady=10)

        self.add_button_daily_report =self.template.create_button(input_frame, self.localization.get('daily_report_button'), self.show_daily_report_input, "Add.TButton", self.image_references['add_icon'], 0, 0, rowspan=1, pady=1, padx=(10, 10))
        self.add_button_minutes =self.template.create_button(input_frame, self.localization.get('meeting_minutes_button'), self.show_meeting_minutes_input, "Add.TButton", self.image_references['add_icon'], 0, 0, rowspan=1, pady=1, padx=(100, 200))
        self.add_button_businesstrip =self.template.create_button(input_frame, self.localization.get('businesstrip_button'), self.show_businesstrip_input, "Add.TButton", self.image_references['add_icon'], 0, 0, rowspan=1, pady=1, padx=(200, 10))
        self.add_button_daily_report.grid(columnspan=2,sticky="w")
        self.add_button_minutes.grid(columnspan=2)
        self.add_button_businesstrip.grid(columnspan=2)

        self.boxno_label, self.subject_entry = self.template.create_label_and_entry(input_frame, self.localization.get('boxno_label'), 1, 0, 50, 12)

        self.subject_entry.bind("<KeyRelease>", self.filter_table_data)

        self.reload_button =self.template.create_button(input_frame, self.localization.get('reload_button'), self.reload_data, "Reload.TButton", self.image_references['reload_icon'], 1, 2, rowspan=1, pady=1, padx=(10, 0))
        
        self.delete_button = self.template.create_button(input_frame, self.localization.get('delete_button'), self.delete_data, "Delete.TButton", self.image_references['delete_icon'], 1, 3, rowspan=1, pady=1, padx=(10, 0))
        
        self.register_button =self.template.create_button(input_frame, self.localization.get('register_button'), self.register, "Register.TButton", self.image_references['register_icon'], 1, 4, rowspan=1, pady=1, padx=(10, 0))
        
        self.csv_button =self.template.create_button(input_frame, 'CSV', self.export_to_csv, "CSV.TButton", self.image_references['csv_icon'], 1, 5, rowspan=1, pady=1, padx=(10, 0))

        self.delete_button["state"] = tk.DISABLED
        
        #self.create_input_window()
        #self.create_list_window()
        
    def export_to_csv(self):
        if len(self.subject_entry.get())>11:
            file_path = f"{self.csv_path}{self.subject_entry.get()}.csv"
            self.template.export_to_csv(self.data_table,file_path,f"{self.localization.get('exported_to_csv')}")
        else:
            messagebox.showinfo(f"{self.localization.get('info')}", f"{self.localization.get('no_boxno')}")

    def update_window_positions(self, event=None):
        self.template.fix_window_position(self.login_window, self.root)
        try:
            self.template.fix_window_position(self.register_window, self.root)
        except Exception as error:
            pass
        try:
            self.template.fix_window_position(self.input_window, self.root)
        except Exception as error:
            pass

    def create_register_window(self):
        self.register_window = tk.Toplevel(self.root)
        self.register_window.title(f"{self.localization.get('register_window_title')}")

        # Create a frame to hold the Listbox and Scrollbar
        frame = ttk.Frame(self.register_window)
        frame.pack(padx=10, pady=10)

        # Create labels and entry widgets
        self.register_username_label, self.register_username_entry = self.template.create_label_and_entry(frame, self.localization.get('user_name'), 0, 0, 110, 11)
        self.register_username_entry.bind("<KeyRelease>", self.filter_register_data)

        self.register_useremail_label, self.register_useremail_entry = self.template.create_label_and_entry(frame, self.localization.get('user_email'), 1, 0, 110, 11)
        self.register_userdepartment_label, self.register_userdepartment_entry = self.template.create_label_and_entry(frame, self.localization.get('user_department'), 2, 0, 110, 11)
        self.register_recipient_label, self.register_recipient_entry = self.template.create_label_and_entry(frame, self.localization.get('recipient'), 3, 0, 110, 11)

        self.register_username_entry.grid(row=0,column=1,columnspan=3)
        self.register_useremail_entry.grid(row=1,column=1,columnspan=3)
        self.register_userdepartment_entry.grid(row=2,column=1,columnspan=3)
        self.register_recipient_entry.grid(row=3,column=1,columnspan=3)

        label_text = self.localization.get('header_meeting_minutes')
        label = ttk.Label(frame, text=label_text, font=("Helvetica", 9), justify="right")
        label.grid(row=4, column=1, padx=10, pady=10)

        frame = frame 
        height, width, fontsize = 12, 70, 8
        row, column = 5, 1  
        self.register_header_minutes_entry = self.template.create_text(frame, height, width, fontsize, row, column)
        self.register_header_minutes_entry.grid(pady=1)

        label_text = self.localization.get('header_daily_report')
        label = ttk.Label(frame, text=label_text, font=("Helvetica", 9), justify="right")
        label.grid(row=4, column=3, padx=10, pady=10)

        frame = frame 
        height, width, fontsize = 12, 70, 8
        row, column = 5, 3  
        self.register_header_daily_report_entry = self.template.create_text(frame, height, width, fontsize, row, column)
        self.register_header_daily_report_entry.grid(pady=1)


        # set the default values for the daily report header
        values = f"""【機密区分：秘/Confidential】

         開示範囲（Disclosure）：モバイルテスト技術部内
         データオーナー（Information Owner）：有馬部長
         お疲れ様です。

         今日の日報を送付致します。"""
        self.register_header_daily_report_entry.insert("end", values + "\n", "formatted_text")

        # set the default values for the meeting minutes header
        values = f"""【機密区分：秘/Confidential】
         開示範囲（Disclosure）：モバイルテスト技術部内
         ====================================================="""
        self.register_header_minutes_entry.insert("end", values + "\n", "formatted_text")

        # Create buttons
        self.template.create_button(frame, self.localization.get('save_button'), self.register_data, "REGISTER.TButton", self.image_references['master_icon'], 0, 4, rowspan=4, pady=(10, 0), padx=(10, 0))
        
        self.register_delete_button = self.template.create_button(frame, self.localization.get('delete_button'), self.delete_register_data, "Delete.TButton", self.image_references['delete_icon'], 0, 4, rowspan=4, pady=(60, 0), padx=(10, 0))

        self.template.create_button(frame, self.localization.get('refresh_button'), self.reset_register_item, "RELOAD.TButton", self.image_references['reload_icon'], 0, 4, rowspan=4, pady=(120, 0), padx=(10, 0), sticky="n")

        # Create the restoreListTable
        columns = (
            self.localization.get('user_name'),
            self.localization.get('user_email'),
            self.localization.get('user_department'),
            self.localization.get('recipient'),
            'id',
            self.localization.get('header_meeting_minutes'),
            self.localization.get('header_daily_report'),
            )
        column_widths = (150, 200,250,400,0,0,0)

        self.register_table = self.template.create_table(frame,columns,column_widths,30,8,8,0,6,self.get_register_table_data)

        # Set the window size and position for the list window
        self.template.set_window_position(root,self.register_window,1200,600)

        # Bind the update_window_positions method to the Configure event of the root window
        self.root.bind("<Configure>", self.update_window_positions)

        # Set the login_window as the master for the register_window
        self.register_window.transient(self.root)

        # Lift the register window above the login window
        self.register_window.lift()

        self.register_username_entry.focus_set()

    def register(self):
        self.create_register_window()
        self.reload_register_data()

    def get_register_table_data(self,event):
        selected_item = self.register_table.selection()
        if selected_item:

            self.register_username_entry.delete(0, tk.END)
            self.register_useremail_entry.delete(0, tk.END)
            self.register_userdepartment_entry.delete(0, tk.END)
            self.register_recipient_entry.delete(0, tk.END)

            item = self.register_table.item(selected_item)
            values = item['values']
            self.register_data_id = values[4]
            self.register_username_entry.insert("end", values[0])
            self.register_useremail_entry.insert("end", values[1])
            self.register_userdepartment_entry.insert("end", values[2])
            self.register_recipient_entry.insert("end", values[3])

            if values[5] != 'None':
                self.register_header_minutes_entry.delete("1.0", "end")
                self.register_header_minutes_entry.insert("end", values[5] + "\n")
            if values[6]!='None':
                self.register_header_daily_report_entry.delete("1.0", "end")
                self.register_header_daily_report_entry.insert("end", values[6] + "\n")

            self.register_delete_button["state"] = tk.NORMAL 
        
    def reset_register_item(self):
        self.reload_register_data()

    def register_data(self):
        run = Models(self.localization,self.globalid)
        result = run.add_register_data(
            self.register_data_id,
            self.globalid,
            self.register_username_entry.get(),
            self.register_useremail_entry.get(),
            self.register_userdepartment_entry.get(),
            self.register_recipient_entry.get(),
            self.register_header_minutes_entry.get("1.0", "end-1c"),
            self.register_header_daily_report_entry.get("1.0", "end-1c"),
            self.pc_name
            )
        if result == 'OK':
            if self.register_data_id == 0:
                messagebox.showinfo(self.localization.get('info'), self.localization.get('created'))
                globalid = self.login_entry.get()
                self.Models = Models(self.localization,globalid)
                user = self.Models.getUser()
                self.create_main_window_item(user)
            else:
                messagebox.showinfo(self.localization.get('info'), self.localization.get('updated'))
            self.reload_register_data()
        else:
            messagebox.showinfo(self.localization.get('info'), "すでに登録済みです。再登録はできません。")

    def filter_register_data(self,event):
        self.register_table.delete(*self.register_table.get_children())

        prefix = self.register_username_entry.get()
        run = Models(self.localization,self.globalid)
        datas = run.retrieve_user_data()
        suggestions = self.template.filter_data(datas,prefix)
        #suggestions = [data for data in datas if data[0].startswith(prefix)]
        for index, suggestion in enumerate(suggestions):
            if index % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
            self.register_table.insert("", "end", values=(suggestion[0], suggestion[1], suggestion[2],suggestion[3],suggestion[4]), tags=(tag, "xfont"))

    def delete_register_data(self):
        confirmation = messagebox.askquestion(self.localization.get('confirm'), f"{self.localization.get('data_delete_confirm')}")
        if confirmation == "yes":
            run = Models(self.localization,self.globalid)
            result = run.delete_register_data(
                self.register_data_id,
                self.register_username_entry.get(),
                self.register_useremail_entry.get(),
                self.register_userdepartment_entry.get()
                )
            if result == 'OK':
                messagebox.showinfo(self.localization.get('info'), self.localization.get('deleted'))
                self.reload_register_data()
        else:
            messagebox.showinfo(self.localization.get('info'), self.localization.get('canceled'))

    def reload_register_data(self):
        run = Models(self.localization,self.globalid)
        data = run.retrieve_user_data()
        
        self.register_username_entry.delete(0, tk.END)
        self.register_useremail_entry.delete(0, tk.END)
        self.register_userdepartment_entry.delete(0, tk.END)
        self.register_recipient_entry.delete(0, tk.END)
        self.register_delete_button["state"] = tk.DISABLED
        self.register_data_id = 0

        self.clear_register_table_data()

        if data:  # Check if data is not empty
            # Configure a tag for the font size
            self.register_table.tag_configure("xfont", font=("Helvetica", 10)) 
            # Insert the new data
            for index, row in enumerate(data):
                if index % 2 == 0:
                    tag = "evenrow"
                else:
                    tag = "oddrow"
                self.register_table.insert("", "end", values=row, tags=(tag, "xfont"))
        else:
            self.register_delete_button["state"] = tk.DISABLED  # Disable the button
            #messagebox.showinfo(self.localization.get('info'), self.localization.get('no_data'))

    def clear_register_table_data(self):
        for item in self.register_table.get_children():
            self.register_table.delete(item)

    def create_login_window(self):
        self.login_window = tk.Toplevel(self.root)
        self.login_window.title(f"{self.localization.get('login_window_title')} {self.pc_name}")

        # Set the window size and position for the list window
        self.template.set_window_position(root,self.login_window,500,100)

        # Create a frame to hold the Listbox and Scrollbar
        frame = ttk.Frame(self.login_window)
        frame.pack(padx=10, pady=10)

        # Create labels and entry widgets
        self.login_label, self.login_entry = self.template.create_label_and_entry(frame, self.localization.get('user_id'), 0, 0, 30, 12)

        self.login_button = ttk.Button(frame, text=f"{self.localization.get('login_button')}", command=self.login)
        self.login_button.grid(row=0, column=2, padx=(0, 100), pady=10)

        self.register_button = ttk.Button(frame, text=f"{self.localization.get('register_button')}", command=self.register)
        self.register_button.grid(row=1, column=0, padx=(10, 0), pady=10)

        self.login_entry.bind('<Return>', self.login_on_enter_pressed)

        self.login_entry.focus_set()
        
        self.login_entry.insert("end", self.globalid)

        if self.globalid:
            self.login()
        else:
            # Call the update_window_positions method to set the initial position
            self.update_window_positions()

            # Bind the update_window_positions method to the Configure event of the root window
            self.root.bind("<Configure>", self.update_window_positions)

            # Set the login_window as the master for the register_window
            self.login_window.transient(self.root)

            # Lift the register window above the login window
            self.login_window.lift()

    def login_on_enter_pressed(self,event):
        self.login()

    def login(self):
        globalid = self.login_entry.get()
        self.Models = Models(self.localization,globalid)
        user = self.Models.getUser()
        if user:
            self.create_main_window_item(user)
        else:
            self.register()
            #messagebox.showerror(f"{self.localization.get('error')}", f"{self.localization.get('user_notfound')}", parent=self.login_window)
            
    def create_main_window_item(self,user):
        self.create_main_section()
        self.create_data_table()
        self.globalid = user[0][0]
        self.user_email= user[0][1]
        self.user_name= user[0][2]
        self.department= user[0][3]
        self.recipient= user[0][4]
        self.login_window.withdraw()
        #message = f" PC: {self.pc_name} USER: {self.globalid} {self.localization.get('login_message')}"
        #self.Models.record_log("LOGIN",message)
        self.root.title(f"{self.title}  PC: {self.pc_name}  【 USER: {self.globalid} 】 ")

        self.reload_data()

part B   def filter_table_data(self, event):
        self.data_table.delete(*self.data_table.get_children())

        prefix = self.subject_entry.get()
        run = Models(self.localization,self.globalid)
        datas = run.retrieve_message_data(self.user_email)
        suggestions = self.template.filter_data(datas,prefix)
        for index, suggestion in enumerate(suggestions):
            if index % 2 == 0:
                tag = "evenrow"
            else:
                tag = "oddrow"
            self.data_table.insert("", "end", values=(suggestion[0], suggestion[1], suggestion[2],suggestion[3],suggestion[4]), tags=(tag, "xfont"))

    def create_input_window(self):
        self.input_window = tk.Toplevel(self.root)
        self.input_window.title(f"{self.localization.get('input_window_title')}")

        # Set the window size and position for the list window
        self.template.set_window_position(root,self.input_window,1000,800)

        # Create a frame to hold the Listbox and Scrollbar
        frame = ttk.Frame(self.input_window)
        frame.pack(padx=10, pady=10)

        self.recipient_label,self.recipient_entry = self.template.create_label_and_text(frame,3,125,10,0,0,self.localization.get('recipient'))
        
        self.subject_label,self.title_entry = self.template.create_label_and_text(frame,1,110,11,1,0,self.localization.get('subject'))
        self.title_entry.configure(bg='#e6e6fa')

        if self.type=='meeting_minutes':
            self.message_header_entry = self.template.create_text(frame,4,110,11,2,1)
            self.message_work_today_entry = self.template.create_text(frame,7,110,11,3,1)
            self.message_comment_entry = self.template.create_text(frame,0,0,11,4,1)
            self.message_work_nextday_entry = self.template.create_text(frame,22,110,11,4,1)
            self.message_work_nextday_entry.configure(bg='#f0f8ff')
        elif self.type=='daily_report':
            self.message_header_entry = self.template.create_text(frame,0,0,11,2,1)
            self.message_work_today_entry = self.template.create_text(frame,11,110,11,2,1)
            self.message_work_nextday_entry = self.template.create_text(frame,11,110,11,3,1)
            self.message_comment_entry = self.template.create_text(frame,11,110,11,4,1)
            self.message_comment_entry.configure(bg='#f0f8ff')
        elif self.type=='businesstrip':
            self.message_header_entry = self.template.create_text(frame,4,110,11,2,1)
            self.message_comment_entry = self.template.create_text(frame,0,0,0,3,1)
            self.message_work_nextday_entry = self.template.create_text(frame,0,0,0,3,1)
            self.message_work_today_entry = self.template.create_text(frame,28,110,11,3,1)
            self.message_work_today_entry.configure(bg='#f0f8ff')

        

        self.template.create_button(frame, self.localization.get('preview_button'), self.send_to_myself_preview, "Reload.TButton", self.image_references['preview_icon'], 7, 1, rowspan=1, pady=20, padx=(0, 500))

        self.template.create_button(frame, self.localization.get('temporary_save_button'), self.temporary_save, "Save.TButton", self.image_references['save_icon'], 7, 1, rowspan=1, pady=20, padx=(300, 300))


        self.template.create_button(frame, self.localization.get('send_button'), self.send_to_recipient, "Delete.TButton", self.image_references['send_icon'], 7, 1, rowspan=1, pady=20, padx=(600, 0))

        
        self.update_window_positions()
         # Bind the update_window_positions method to the Configure event of the root window
        self.root.bind("<Configure>", self.update_window_positions)

        # Set the login_window as the master for the register_window
        self.input_window.transient(self.root)

        # Lift the register window above the login window
        self.input_window.lift()

        self.input_window.withdraw()

    def show_daily_report_input(self):
        self.type = 'daily_report'
        self.data_id = 0
        self.show_input()

    def show_meeting_minutes_input(self):
        self.type = 'meeting_minutes'
        self.data_id = 0
        self.show_input()

    def show_businesstrip_input(self):
        self.type = 'businesstrip'
        self.data_id = 0
        self.show_input()

    def show_input(self):
        try:
            self.input_window.withdraw()
        except Exception as error:
            pass
        self.create_input_window()
        self.input_window.deiconify()
        self.input_window.lift()
        if self.data_id <= 0:
            self.input_window.title(f"{self.localization.get('input_window_title')} 新規入力 ({self.user_email})")

            globalid = self.login_entry.get()
            self.Models = Models(self.localization,globalid)
            user = self.Models.getUser()

            today = datetime.now().strftime("%Y%m%d")
            year = datetime.now().strftime("%Y")
            month = datetime.now().strftime("%m")
            day = datetime.now().strftime("%d")

            if self.data_id==0:
                self.recipient= user[0][4]
                self.recipient_entry.insert("end", self.recipient + "\n")
                if self.type=='meeting_minutes':
                    self.title_entry.insert("end",f"秘_【議事録】{self.user_name}_{today}（{self.template.day_of_week_japanese(datetime.now())}）")
                    self.Models = Models(self.localization,self.globalid)
                    user = self.Models.getUser()
                    self.message_header_entry.insert("end",f"{user[0][6]}\n表記の件、議事録は下記に参考送付いたします。\n")
                    self.message_work_today_entry.insert("end",f"""日付: {year}年{month}月{day}日（{self.template.day_of_week_japanese(datetime.now())}）\n出席者: {self.user_name} \n\n議題:""")
                    self.message_work_nextday_entry.insert("end",f"【議事概要】 \n")
                    self.message_comment_entry.insert("end","")
                elif self.type=='daily_report':
                    user = self.Models.getUser()
                    self.message_header_entry.insert("end",f"{user[0][5]}\n\n")
                    self.title_entry.insert("end",f"秘_【日報】{self.user_name}_{today}（{self.template.day_of_week_japanese(datetime.now())}）")
                    self.message_work_today_entry.insert("end",f"【業務内容】今日\n")
                    self.message_work_nextday_entry.insert("end",f"【業務内容】翌日\n")
                    self.message_comment_entry.insert("end",f"【所感】\n")
                elif self.type=='businesstrip':
                    user = self.Models.getUser()
                    self.message_header_entry.insert("end",f"{user[0][6]}\n表記の件、出張連絡は下記に参考送付いたします。\n")
                    self.title_entry.insert("end",f"秘_【出張連絡】{self.user_name}_{today}（{self.template.day_of_week_japanese(datetime.now())}）")
                    self.message_work_today_entry.insert("end",f"""下記日程で博多オフィスへ出張いたします。
                                                        \n【日程】\n {year}年{month}月{day}日（{self.template.day_of_week_japanese(datetime.now())}）
                                                        \n\n【目的】
                                                        \n\n【交通手段】
                                                        \n\n【宿泊地】
                                                        \n\n【同行者】\n なし
                                                        \n\n 以上、よろしくお願いします。
                                                        """)
                    self.message_work_nextday_entry.insert("end",f"")
                    self.message_comment_entry.insert("end",f"")

            self.message_work_today_entry.focus_set()
        else:
            self.input_window.title(f"{self.localization.get('input_window_title')}  ID-{self.data_id} 編集 ({self.sender})" )
    
    # send the email for myself for review
    def send_to_myself_preview(self):
        body = self.send_message()
        sender = {'sender':self.user_email}
        self.email_message.update(sender)
        recipient = {'recipient':self.user_email}
        self.email_message.update(recipient)
        subject = {'subject':f'{self.title_entry.get("1.0", "end-1c")}'}
        self.email_message.update(subject)
        message_body = {'body':body}
        self.email_message.update(message_body)

        recipient = self.recipient_entry.get("1.0", "end-1c")
        sender = self.user_email
        subject = self.title_entry.get("1.0", "end-1c")
        message_today = self.message_work_today_entry.get("1.0", "end-1c")
        message_nexday = self.message_work_nextday_entry.get("1.0", "end-1c")
        message_comment = self.message_comment_entry.get("1.0", "end-1c")
        header = self.message_header_entry.get("1.0", "end-1c")
        if len(subject) > 3:
            run = Models(self.localization,self.globalid)
            result = run.update_preview_data(
            self.data_id,
            recipient,
            sender,
            subject,
            message_today,
            message_nexday,
            message_comment,
            self.pc_name,
            header
            )
        if result == 'OK':
            self.template.send_email(self.email_message)
            messagebox.showinfo(f"{self.localization.get('info')}", f"{self.localization.get('check_your_email')}", parent=self.input_window)
            self.reload_data()
            self.input_window.withdraw()
        else:
            messagebox.showinfo(f"{self.localization.get('info')}", result, parent=self.input_window)
        
    # send the email to the public as per the list in the recipients
    def send_to_recipient(self):
        body = self.send_message()
        sender = {'sender':self.user_email}
        self.email_message.update(sender)
        recipient = {'recipient':f'{self.recipient_entry.get("1.0", "end-1c")}'}
        self.email_message.update(recipient)
        subject = {'subject':f'{self.title_entry.get("1.0", "end-1c")}'}
        self.email_message.update(subject)
        message_body = {'body':body}
        self.email_message.update(message_body)

        if self.send == 1:
            message = f"{self.localization.get('data_resend_confirm')}"
        else:
            message = f"{self.localization.get('data_send_confirm')}"

        confirmation = messagebox.askquestion(self.localization.get('confirm'), message)
        if confirmation == "yes":
            self.template.send_email(self.email_message)
            recipient = self.recipient_entry.get("1.0", "end-1c")
            sender = self.user_email
            subject = self.title_entry.get("1.0", "end-1c")
            message_today = self.message_work_today_entry.get("1.0", "end-1c")
            message_nexday = self.message_work_nextday_entry.get("1.0", "end-1c")
            message_comment = self.message_comment_entry.get("1.0", "end-1c")
            header = self.message_header_entry.get("1.0", "end-1c")
            if len(subject) > 3:
                run = Models(self.localization,self.globalid)
                result = run.update_send_data(
                self.data_id,
                recipient,
                sender,
                subject,
                message_today,
                message_nexday,
                message_comment,
                self.pc_name,
                header
                )
            if result == 'OK':
                messagebox.showinfo(self.localization.get('info'), self.localization.get('data_send'))
                self.reload_data()
                self.input_window.withdraw()
        else:
            messagebox.showinfo(self.localization.get('info'), self.localization.get('canceled'))

    part C    def temporary_save(self):
        recipient = self.recipient_entry.get("1.0", "end-1c")
        sender = self.user_email
        subject = self.title_entry.get("1.0", "end-1c")
        message_today = self.message_work_today_entry.get("1.0", "end-1c")
        message_nexday = self.message_work_nextday_entry.get("1.0", "end-1c")
        message_comment = self.message_comment_entry.get("1.0", "end-1c")
        header = self.message_header_entry.get("1.0", "end-1c")
        if len(subject) > 3:
            run = Models(self.localization,self.globalid)
            result = run.update_preview_data(
            self.data_id,
            recipient,
            sender,
            subject, 
            message_today,
            message_nexday,
            message_comment,
            self.pc_name,
            header
            )
        if result == 'OK':
            messagebox.showinfo(f"{self.localization.get('info')}", f"{self.localization.get('saved')}", parent=self.input_window)
            self.reload_data()
            self.input_window.withdraw()

    def send_message(self):
        self.email_message = {
                "recipient": "",
                "recipient_cc": "",
                "subject": "",
                "header": "",
                "body": "",
                "footer":"",
                "sender" : self.sender,
                "smtp_server" : 'xxx.co.jp',
                "smtp_port" : 25
            }
        
        message_header_lines = "<br>".join(self.message_header_entry.get("1.0", "end-1c").splitlines())
        #message_work_today = self.message_work_today_entry.get("1.0", "end-1c").replace('\n', '<br>')
        message_work_today_lines = self.message_work_today_entry.get("1.0", "end-1c").splitlines()
        # Get the first line
        if message_work_today_lines:
            first_line_message_work_today = message_work_today_lines[0]
        else:
            first_line_message_work_today = ""
        # Join the remaining lines (excluding the first one)
        remaining_lines_message_work_today = "<br>".join(message_work_today_lines[1:])

        #message_work_nextday = self.message_work_nextday_entry.get("1.0", "end-1c").replace('\n', '<br>')
        message_work_nextday_lines = self.message_work_nextday_entry.get("1.0", "end-1c").splitlines()
        # Get the first line
        if message_work_nextday_lines:
            first_line_message_work_nextday = message_work_nextday_lines[0]
        else:
            first_line_message_work_nextday = ""
        # Join the remaining lines (excluding the first one)
        remaining_lines_message_work_nextday = "<br>".join(message_work_nextday_lines[1:])

        message_comment_lines = self.message_comment_entry.get("1.0", "end-1c").splitlines()
        # Get the first line
        if message_comment_lines:
            first_line_comment = message_comment_lines[0]
        else:
            first_line_comment = ""
        # Join the remaining lines (excluding the first one)
        remaining_lines_comment = "<br>".join(message_comment_lines[1:])

        self.Models = Models(self.localization,self.globalid)
        user = self.Models.getUser()
        body = f"""
        <table style="font-family: 'Noto Sans CJK JP', sans-serif;font-size: 14px;">
        <tr><td>
        {message_header_lines}
         </td></tr></table>
         <br>
        """
        if len(first_line_message_work_today)>1:
            body += f"""
            <table style="border-collapse: collapse; border: 1px solid black;font-family: 'Noto Sans CJK JP', sans-serif;font-size: 14px;font-size: 14px;">
            <tr>
                <th style="background:#e6e6fa;padding: 8px; text-align: left; border: 1px solid black; width:1000px;">{first_line_message_work_today}</th>
            </tr>
            <tr>
            <td style="background:#FFFFFF;padding: 8px; text-align: left; border: 1px solid black;height: 100%;">{remaining_lines_message_work_today}</td>
            </tr>
            </table><br>
            """

        if len(first_line_message_work_nextday)>1:
            body += f"""
            <table style="border-collapse: collapse; border: 1px solid black;font-family: 'Noto Sans CJK JP', sans-serif;font-size: 14px;">
            <tr>
                <th style="background:#e6e6fa;padding: 8px; text-align: left; border: 1px solid black; width:1000px;">{first_line_message_work_nextday}</th>
            </tr>
            <tr>
            <td style="background:#FFFFFF;padding: 8px; text-align: left; border: 1px solid black;height: 100%;">{remaining_lines_message_work_nextday}</td>
            </tr>
            </table><br>
            """

        if len(first_line_comment)>1:
            body += f"""
            <table style="border-collapse: collapse; border: 1px solid black;width=500px;font-family: 'Noto Sans CJK JP', sans-serif;font-size: 14px;">
            <tr>
                <th style="background:#e6e6fa;padding: 8px; text-align: left; border: 1px solid black; width:1000px;">{first_line_comment}</th>
            </tr>
            <tr>
            <td style="background:#FFFFFF;padding: 8px; text-align: left; border: 1px solid black;height: 100%;">{remaining_lines_comment}</td>
            </tr>
            </table><br>
            """

        body += f"""        
        <br><br>
        <table style="font-family: 'Noto Sans CJK JP', sans-serif;font-size: 14px;">
        <tr><td>
        以上、よろしくお願いいたします。
        </td></tr></table>
        <br><br>
        
        <table style="width:500px;border-collapse: collapse; border: 1px solid blue;font-family: 'Noto Sans CJK JP', sans-serif;font-size: 14px;">
        <tr><td>
        　ソニーセミコンダクタマニュファクチャリング株式会社　　<br>
        　{self.department}　<br>
        　{self.user_name}　<br>
        　Mail: {self.user_email}　<br>
        </td></tr>
        </table>


        """

        return body

    def create_data_table(self):
        columns = (
            self.localization.get('send_date'), 
            self.localization.get('subject'), 
            self.localization.get('send'),
            self.localization.get('recipient'),
            "","","","")
        self.column_widths = (130, 500, 100,500,0,0,0,0)
        self.data_table = self.template.create_table_pack(
            self.root,
            columns,
            self.column_widths,
            30, # row height
            8,  # font size
            self.get_table_data, # TreeViewSelect
            self.edit_table_data # DoubleClick
        )

        # Hide column
        self.template.hide_column(self.data_table, 5)
        self.template.hide_column(self.data_table, 6)
        self.template.hide_column(self.data_table, 7)
        self.template.hide_column(self.data_table, 8)

        self.data_table.heading("#0", text="全削除", anchor="center", command=self.delete_all_data) 
        self.data_table.column("#0", width=70, anchor="w")


    def delete_all_data(self):
        confirmation = messagebox.askquestion(self.localization.get('confirm'), f"{self.localization.get('all_data_delete_confirm')}")
        if confirmation == "yes":
            run = Models(self.localization,self.globalid)
            result = run.delete_data_all(self.user_email)
            if result == 'OK':
                self.reload_data()
                messagebox.showinfo(self.localization.get('info'), self.localization.get('deleted'))
        else:
            messagebox.showinfo(self.localization.get('info'), self.localization.get('canceled'))

    def edit_table_data(self, event):
        selected_item = self.data_table.selection()

        if selected_item:
            item = self.data_table.item(selected_item)
            values = item['values']

            self.data_id, self.sender, self.send = values[7], values[8], values[2]
            self.set_value_to_show_table(values)
            

    def set_value_to_show_table(self,values):
        if self.localization.get('meeting_minutes_button') in values[1]:
            self.type = 'meeting_minutes'
        elif self.localization.get('daily_report_button') in values[1]:
            self.type = 'daily_report'

        self.show_input()
        self.recipient_entry.insert("end", f"{values[3]}\n")
        self.title_entry.insert("end", values[1])

        self.message_work_today_entry.insert("end", f"{values[4]}\n")
        self.message_work_nextday_entry.insert("end", f"{values[5]}\n")
        self.message_comment_entry.insert("end", f"{values[6]}\n")

        user = self.Models.getUser()
        header_value = ""

        if values[9] != "None" and self.type != 'daily_report':
            header_value = f"{values[9]}\n"
        elif self.type == 'daily_report':
            header_value = f"{user[0][5]}\n"
        else:
            if self.type == 'meeting_minutes':
                header_value = f"{user[0][6]}\n表記の件、議事録は下記に参考送付いたします。\n"
            elif self.type == 'daily_report':
                header_value = f"{user[0][5]}\n\n"
            elif self.type == 'businesstrip':
                header_value = f"{user[0][6]}\n表記の件、出張連絡は下記に参考送付いたします。\n"

        self.message_header_entry.insert("end", header_value + "\n")

    def get_table_data(self,event):
        if self.template.selected_column != '0':
            selected_item = self.data_table.selection()
            if selected_item:
                self.subject_entry.delete(0, tk.END)
                item = self.data_table.item(selected_item)
                values = item['values']
                self.data_id = values[7]
                self.sender = values[8]
                self.send = values[2]
                self.delete_button["state"] = tk.NORMAL 
                self.subject_entry.insert(0, values[1])
        else:
            selected_item = self.data_table.selection()
            if selected_item:
                item = self.data_table.item(selected_item)
                values = item['values']

                self.data_id, self.sender, self.send = -1, values[8], values[2]
                self.set_value_to_show_table(values)

           
    def reload_data(self):
        run = Models(self.localization,self.globalid)
        data = run.retrieve_message_data(self.user_email)
        
        self.subject_entry.delete(0, tk.END)
        self.add_button_minutes["state"] = tk.NORMAL
        self.add_button_daily_report["state"] = tk.NORMAL
        self.delete_button["state"] = tk.DISABLED
        self.data_id = 0
        self.send = 0

        self.clear_table_data()

        if data:  # Check if data is not empty
            # Configure a tag for the font size
            self.data_table.tag_configure("xfont", font=("Helvetica", 10)) 
            # Insert the new data
            for index, row in enumerate(data):
                if index % 2 == 0:
                    tag = "evenrow"
                else:
                    tag = "oddrow"
                # Apply "xfont" tag to all columns
                item_id = self.data_table.insert("", "end", text="COPY", values=row, tags=(tag, "xfont"))


        else:
            self.delete_button["state"] = tk.DISABLED  # Disable the button
            #messagebox.showinfo(self.localization.get('info'), self.localization.get('no_data'))


    def clear_table_data(self):
        for item in self.data_table.get_children():
            self.data_table.delete(item)

    def delete_data(self):
        confirmation = messagebox.askquestion(self.localization.get('confirm'), f"{self.localization.get('data_delete_confirm')}")
        if confirmation == "yes":
            run = Models(self.localization,self.globalid)
            result = run.delete_data(self.data_id,self.subject_entry.get())
            if result == 'OK':
                self.reload_data()
                messagebox.showinfo(self.localization.get('info'), self.localization.get('deleted'))
        else:
            messagebox.showinfo(self.localization.get('info'), self.localization.get('canceled'))

   part D   class Models:
    def __init__(self,localization,globalid):

        self.globalid = globalid
        self.template = template()
        self.localization = localization

        self.db_params = {
            'database' : 'office',
            'host' : '43.24.188.114',
            'user' : 'office',
            'password' : 'office',
            'port' : '5432',
            'query': None,
            'values': None
        }
    
    def getUser(self):
        query = f"""SELECT 
        user_id,
        user_email,
        user_name,
        department,
        recipient,
        header_daily_report,
        header_meeting_minutes 
        FROM office.tm_user WHERE user_id = '{self.globalid}' ORDER BY user_name"""  
        query = {'query':query}
        self.db_params.update(query)
        return self.template.db_fetchall(self.db_params)

    def get_backuped_boxno(self):
        query = f"SELECT DISTINCT BOX_NO FROM backup_lot_tbl "
        query = {'query':query}
        self.db_params.update(query)
        result = self.template.mySQL_fetchall(self.db_params)
        return result

    def delete_data(self,id,subject):
        query = f"""DELETE FROM office.td_daily_report WHERE id = '{id}'"""
        query = {'query':query}
        self.db_params.update(query)
        result = self.template.db_commit(self.db_params)

        message = f" USER: {self.globalid}  {subject} {self.localization.get('deleted')}"
        self.template.logging_date(message)
        return result 
    
    def delete_data_all(self,id):
        query = f"""DELETE FROM office.td_daily_report WHERE sender = '{id}'"""
        query = {'query':query}
        self.db_params.update(query)
        result = self.template.db_commit(self.db_params)

        message = f" USER: {self.globalid}   {self.localization.get('all_deleted')}"
        self.template.logging_date(message)
        return result 
    
    def delete_register_data(self,id,username,useremail,userdepartment):
        
        query = f"""DELETE FROM office.tm_user WHERE user_id = '{id}'"""
        query = {'query':query}
        self.db_params.update(query)
        result = self.template.db_commit(self.db_params)

        message = f" Department: {userdepartment} USER: {username} : {useremail} 　{self.localization.get('deleted')}"
        self.record_log("CREATE",message)
        
        return result 


    def update_send_data(self,id,recipient,sender,subject,message_today,message_nextday,message_comment,pc_name,header):
        if id == 0:
            query = f"""
                INSERT INTO office.td_daily_report (recipient,sender,subject,message_today,message_nextday,message_comment,send,header) 
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """
            query = {'query':query}
            self.db_params.update(query)
            values = (recipient, sender, subject, message_today, message_nextday, message_comment, 1, header)
            values = {'values':values}
            self.db_params.update(values)
            result = self.template.db_commit_values(self.db_params)
            message = f" PC: {pc_name} USER: {sender} : {subject} 　{self.localization.get('created')}"
            self.record_log("CREATE",message)
        else:
            query = f"""
                UPDATE office.td_daily_report SET
                recipient=%s,
                sender=%s, 
                subject=%s,
                header=%s,
                message_today=%s,
                message_nextday=%s,
                message_comment=%s,
                send=%s 
                WHERE id = '{id}' """
            query = {'query':query}
            self.db_params.update(query)
            values = (recipient, sender, subject, header, message_today, message_nextday, message_comment, 1)
            values = {'values':values}
            self.db_params.update(values)
            result = self.template.db_commit_values(self.db_params)
            message = f" PC: {pc_name} USER: {sender} : {subject} 　{self.localization.get('data_send')}"
            self.record_log("UPDATE",message)
        
        return result 

    def update_preview_data(self,id,recipient,sender,subject,message_today,message_nextday,message_comment,pc_name,header):
        if id == 0:
            query = "INSERT INTO office.td_daily_report (recipient, sender, subject, message_today, message_nextday, message_comment, header, send) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"
            query = {'query':query}
            self.db_params.update(query)
            
            # Provide values as a tuple as the second parameter of the execute method
            values = (recipient, sender, subject, message_today, message_nextday, message_comment, header, 0)
            values = {'values':values}
            self.db_params.update(values)

            result = self.template.db_commit_values(self.db_params)
            message = f" PC: {pc_name} USER: {sender} : {subject} 　{self.localization.get('created')}"
            self.record_log("CREATE",message)
        else:
            query = f"""
                UPDATE office.td_daily_report SET
                recipient=%s,
                sender=%s, 
                subject=%s,
                header=%s,
                message_today=%s,
                message_nextday=%s,
                message_comment=%s
                WHERE id = '{id}' """
            query = {'query':query}
            self.db_params.update(query)
            values = (recipient, sender, subject, header, message_today, message_nextday, message_comment)
            values = {'values':values}
            self.db_params.update(values)
            result = self.template.db_commit_values(self.db_params)
            message = f" PC: {pc_name} USER: {sender} : {subject} 　{self.localization.get('updated')}"
            self.record_log("UPDATE",message)
        
        return result 

    def add_register_data(self,id,userid,name,email,department,recipient,header_minutes,header_daily_report,pc):
        if id == 0:
            query = f"""
                INSERT INTO office.tm_user (user_id,user_email,user_name,pc_name,department,recipient,header_daily_report,header_meeting_minutes) 
                VALUES(%s,%s,%s,%s,%s,%s,%s,%s)"""
            query = {'query':query}
            self.db_params.update(query)
            values = (userid, email, name, pc, department, recipient, header_daily_report, header_minutes)
            values = {'values':values}
            self.db_params.update(values)
            result = self.template.db_commit_values(self.db_params)
            message = f" PC: {pc} USER: {name} : {userid} {self.localization.get('created')}"
            self.record_log("CREATE",message)
        else:
            userid = id
            query = f"""
                UPDATE office.tm_user SET 
                pc_name=%s,
                user_email=%s, 
                user_name=%s,
                department=%s,
                header_daily_report=%s ,
                header_meeting_minutes=%s ,
                recipient=%s 
                WHERE user_id = '{userid}' """
            query = {'query':query}
            self.db_params.update(query)
            values = (pc, email, name, department, header_daily_report, header_minutes, recipient)
            values = {'values':values}
            self.db_params.update(values)
            result = self.template.db_commit_values(self.db_params)
            message = f" PC: {pc} USER: {name} : {userid} {self.localization.get('updated')}"
            self.record_log("UPDATE",message)
        
        return result 

    def retrieve_message_data(self, sender):
        base_query = (
            "SELECT create_datetime, subject, send, recipient, "
            "message_today, message_nextday, message_comment, id, "
            "sender, header FROM office.td_daily_report"
        )

        query = f"{base_query} WHERE sender = '{sender}' ORDER BY sender, create_datetime desc, subject"

        if self.globalid == '0020H34707':
            query = f"{base_query} ORDER BY sender, create_datetime, subject"

        query = {'query':query}
        self.db_params.update(query)
        return self.template.db_fetchall(self.db_params)

    
    # def retrieve_user_data(self):
    #     query = f"""select user_name,user_email,department,recipient,user_id,header_meeting_minutes,header_daily_report
    #     FROM  office.tm_user WHERE user_id = '{self.globalid}' ORDER BY user_id"""
    #     if self.globalid == '0020H34707':
    #         query = f"""select user_name,user_email,department,recipient,user_id,header_meeting_minutes,header_daily_report
    #         FROM  office.tm_user"""
    #     query = {'query':query}
    #     self.db_params.update(query)
    #     return self.template.db_fetchall(self.db_params)
    def retrieve_user_data(self):
        base_query = (
            "SELECT user_name, user_email, department, recipient, "
            "user_id, header_meeting_minutes, header_daily_report "
            "FROM office.tm_user"
        )

        query = f"{base_query} WHERE user_id = '{self.globalid}' ORDER BY user_name"

        if self.globalid == '0020H34707':
            query = base_query

        query = {'query':query}
        self.db_params.update(query)
        return self.template.db_fetchall(self.db_params)


    def record_log(self,type,log):
        query = f"INSERT INTO office.td_log (type,log) VALUES{type,log}"
        query = {'query':query}
        self.db_params.update(query)
        self.template.db_commit(self.db_params)

if __name__ == "__main__":    
    root = tk.Tk()
    app = Views(root)
    root.mainloop()


