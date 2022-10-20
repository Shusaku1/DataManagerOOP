# tkinterのインポート
import tkinter as tk
import tkinter.constants
import tkinter.ttk as ttk

from tkinter import messagebox, filedialog
from tkinter.constants import MULTIPLE

from tkinterdnd2 import DND_FILES
import pandas as pd
from tkinterdnd2 import TkinterDnD


class DataManagement(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.button_to_csv = None
        self.button_to_excel = None
        self.button_show_df = None
        self.file_names_listbox2 = None
        self.df2_column_listbox = None
        self.df1_column_listbox = None
        self.df_for_join = []
        self.tree = None
        self.column_names = []
        self.selected = []
        self.df_selected = None
        self.df_CONCAT = None
        self.cell_count = None
        self.columns_count = None
        self.rows_count = None
        self.df = None
        self.master = master

        self.master.geometry("1400x800")
        self.master.title("Tkinter with Class Template")
        self.master.resizable(None)

        # フレームの作成と設置
        self.frame_dropbox = ttk.Frame(self.master)
        self.frame_dropbox.grid(column=0, row=0, sticky=tk.NSEW, padx=30, pady=30)

        self.frame1 = ttk.Frame(self.master)
        self.frame1.grid(column=1, row=0, sticky=tk.NSEW, padx=5, pady=30)

        self.frame_space1 = ttk.Frame(self.master)
        self.frame_space1.grid(column=0, row=1, sticky=tk.NSEW, padx=5, pady=10)

        self.frame2 = ttk.Frame(self.master)
        self.frame2.grid(column=0, row=2, sticky=tk.NSEW, padx=5, pady=10)

        self.frame3 = ttk.Frame(self.master)
        self.frame3.grid(column=1, row=2, sticky=tk.NSEW, padx=5, pady=10)

        self.frame4 = ttk.Frame(self.master, width=500, height=70)
        self.frame4.grid(column=2, row=0, rowspan=3, columnspan=1, sticky=tk.NSEW, padx=150, pady=30)

        self.frame4.grid(padx=5, pady=5, ipadx=5, ipady=5)
        # 1列目を可変サイズとする
        self.frame4.columnconfigure(0, weight=1)
        # 1行目を可変サイズとする
        self.frame4.rowconfigure(0, weight=1)
        # 内部のサイズに合わせたフレームサイズとしない
        self.frame4.grid_propagate(False)

        self.frame5 = ttk.Frame(self.master)
        self.frame5.grid(column=2, row=3, sticky=tk.NSEW, padx=5, pady=10)

        # Frame1
        # 各種ウィジェットの作成

        # self.label1 = ttk.Label(self.frame1, text="Excelファイル名：")
        # self.label2 = ttk.Label(self.frame1, text="シート名：")

        # df情報
        self.label_rows = ttk.Label(self.frame1, text="行数:")  # rows
        self.label_columns = ttk.Label(self.frame1, text="列数:")  # columns
        self.label_cells = ttk.Label(self.frame1, text="セル数:")  # cells
        self.label_rowCount = ttk.Label(self.frame1, text="  ")  # rows
        self.label_columnCount = ttk.Label(self.frame1, text="  ")  # columns
        self.label_cellCount = ttk.Label(self.frame1, text="  ")
        self.button_inner_join = ttk.Button(self.frame1, text='INNER JOIN', width=20, command=self.inner_join_window)
        self.button_left_join = ttk.Button(self.frame1, text='LEFT OUTER JOIN', width=20,
                                           command=self.inner_join_window)
        self.button_right_join = ttk.Button(self.frame1, text='RIGHT OUTER JOIN', width=20,
                                            command=self.inner_join_window)

        # 空白作り用
        self.label_blank1 = ttk.Label(self.frame1, text="　　　　　　")
        self.label_blank2 = ttk.Label(self.frame1, text="　　　　　　")
        self.label_blankJoin = ttk.Label(self.frame1, text="       ")
        """
        # 入力とボタン
        self.entry1 = ttk.Entry(self.frame1)
        self.entry2 = ttk.Entry(self.frame1)
        self.button_loading = ttk.Button(self.frame1, text="読み込む", command=self.load_excel_file)

        # 各種ウィジェットの設置　(frame1)
        # df基本情報
        self.label1.grid(row=0, column=0)
        self.entry1.grid(row=0, column=1)

        self.label2.grid(row=1, column=0)
        self.entry2.grid(row=1, column=1)
        """
        self.label_rows.grid(row=0, column=3)
        self.label_columns.grid(row=1, column=3)
        self.label_cells.grid(row=2, column=3)

        self.label_rowCount.grid(row=0, column=4)
        self.label_columnCount.grid(row=1, column=4)
        self.label_cellCount.grid(row=2, column=4)

        self.button_inner_join.grid(row=4, column=3)
        self.button_left_join.grid(row=5, column=3)
        self.button_right_join.grid(row=6, column=3)

        self.label_blank1.grid(row=0, column=2)
        self.label_blank2.grid(row=1, column=2)
        self.label_blankJoin.grid(row=3, column=3)

        # Entryウィジェットへ文字列のセット
        """
        self.entry1.insert(tk.END, "Hello_World")
        self.entry2.insert(tk.END, "Sheet1")
        """
        # Frame_dropbox
        self.file_names_listbox = tk.Listbox(self.frame_dropbox, selectmode=tk.SINGLE, background="darkgray")
        self.file_names_listbox.pack(fill=tk.X)
        self.file_names_listbox.drop_target_register(DND_FILES)
        self.file_names_listbox.dnd_bind("<<Drop>>", self.drop_inside_listbox)
        self.file_names_listbox.bind("<Double-1>", self.select_file)
        self.open_button = ttk.Button(self.frame_dropbox, text='参照', command=self.select_file_dropbox)
        self.open_button.pack(expand=True)

        # Frame space1
        self.label_blank6 = ttk.Label(self.frame_space1, text="　　　　　　　　　　　　")
        self.label_blank7 = ttk.Label(self.frame_space1, text="　　　　　　　　　　　　")
        self.label_blank6.grid(row=0, column=2)
        self.label_blank7.grid(row=1, column=2)

        # Frame2
        # カラムの選択,新しいdf作成用
        self.label_columnList = ttk.Label(self.frame2, text="カラム一覧")
        self.label_blank3 = ttk.Label(self.frame2, text="　　　　　　　　　　")
        self.label_blank4 = ttk.Label(self.frame2, text="　　　　　　　　　　")
        # カラムリスト配置
        self.label_blank3.grid(row=0, column=0)
        self.label_blank4.grid(row=0, column=3)
        self.label_columnList.grid(row=0, column=1)

        # スクロールバーの作成
        self.scrollbar = tk.Scrollbar(self.frame2, orient=tk.VERTICAL, command=tk.Listbox.yview)

        # リストボックス、スクロールバーの配置
        self.listbox = tk.Listbox(self.frame2, width=27, height=15, selectmode=MULTIPLE,
                                  yscrollcommand=self.scrollbar.set)
        self.listbox.grid(row=1, column=1)
        self.scrollbar.config(command=self.listbox.yview)
        self.scrollbar.grid(row=1, column=2, sticky=[tk.N, tk.S])

        # 全選択、全クリア、確定ボタン
        self.button_selectAll = ttk.Button(self.frame2, text="全選択", command=self.select_all)
        self.button_clearAll = ttk.Button(self.frame2, text="全クリア", command=self.clear_all)
        self.button_confirm = ttk.Button(self.frame2, text="確定", command=self.confirm)
        # 各種ボタン配置
        self.button_selectAll.grid(row=2, column=1)
        self.button_clearAll.grid(row=3, column=1)
        self.button_confirm.grid(row=4, column=1)

        # Frame3
        # データ探索 指定カラムから入力されたワードを含む行を探す
        self.label_search = ttk.Label(self.frame3, text="データ探索")
        self.label_search.grid(row=0, column=1)

        self.label_keyWord = ttk.Label(self.frame3, text="検索ワード：")
        self.label_keyWord.grid(row=1, column=0)

        self.entry_keyWord = ttk.Entry(self.frame3, width=27)
        self.entry_keyWord.grid(row=1, column=1)

        self.button_search = ttk.Button(self.frame3, text="探索", command=self.search)
        self.button_search.grid(row=1, column=2)
        # スペース
        self.label_blank5 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank5.grid(row=2, column=0)

        self.label_rowsKeyword = ttk.Label(self.frame3, text="キーワードを含む行数：")
        self.label_rowsKeyword.grid(row=3, column=0)

        self.label_rowsKeywordResult = ttk.Label(self.frame3, text="")
        self.label_rowsKeywordResult.grid(row=3, column=1)

        # スペース
        self.label_blank8 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank8.grid(row=4, column=0)
        self.label_blank9 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank9.grid(row=5, column=0)
        self.label_blank10 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank10.grid(row=6, column=0)
        self.label_blank11 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank11.grid(row=10, column=0)

        # 列操作
        self.label_columnOperation = ttk.Label(self.frame3, text="データの列操作")
        self.label_columnOperation.grid(row=7, column=1)

        # 列追加
        self.label_addColumn = ttk.Label(self.frame3, text="列追加：")
        self.label_addColumn.grid(row=8, column=0)

        self.entry_newColumn = ttk.Entry(self.frame3, width=27)
        self.entry_newColumn.grid(row=8, column=1)
        self.entry_newColumn.insert(tk.END, "追加したい列名を入力")

        self.button_search = ttk.Button(self.frame3, text="追加", command=self.add_column)
        self.button_search.grid(row=8, column=2)

        # 列削除
        self.label_columnDelete = ttk.Label(self.frame3, text="列削除：")
        self.label_columnDelete.grid(row=9, column=0)

        # Adding combobox drop down list
        self.n = tk.StringVar()
        self.column_combobox1 = ttk.Combobox(self.frame3, width=27, textvariable=self.n)

        self.column_combobox1.grid(row=9, column=1)

        self.button_deleteColumn = ttk.Button(self.frame3, text="削除", command=self.delete_column)
        self.button_deleteColumn.grid(row=9, column=2)

        # 列交換
        self.label_columnSwitch = ttk.Label(self.frame3, text="列交換：")
        self.label_columnSwitch.grid(row=11, column=0)

        # Adding combobox drop down list
        self.n = tk.StringVar()
        self.column_combobox2 = ttk.Combobox(self.frame3, width=27, textvariable=self.n)

        self.column_combobox2.grid(row=11, column=1)

        self.m = tk.StringVar()
        self.column_combobox3 = ttk.Combobox(self.frame3, width=27, textvariable=self.m)

        self.column_combobox3.grid(row=12, column=1)

        self.button_searchColumn = ttk.Button(self.frame3, text="交換", command=self.switch_column)
        self.button_searchColumn.grid(row=11, column=2)

        # 行操作
        # スペース
        self.label_blank11 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank11.grid(row=12, column=0)
        self.label_blank12 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank12.grid(row=13, column=0)
        self.label_blank13 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank13.grid(row=14, column=0)
        self.label_blank14 = ttk.Label(self.frame3, text="　　　　　　　　　　　　")
        self.label_blank14.grid(row=18, column=0)

        # 行操作
        self.label_rowOperation = ttk.Label(self.frame3, text="データの行操作")
        self.label_rowOperation.grid(row=15, column=1)

        # 行追加
        self.label_addRow = ttk.Label(self.frame3, text="行追加：")
        self.label_addRow.grid(row=16, column=0)
        self.list_entry = []

        self.button_addRow = ttk.Button(self.frame3, text="追加", width=27, command=self.add_window)
        self.button_addRow.grid(row=16, column=1)

        # self.button_searchRow = ttk.Button(self.frame3, text="追加", command=self.add_row)
        # self.button_searchRow.grid(row=16, column=2)

        # 行削除
        self.label_deleteRow = ttk.Label(self.frame3, text="行削除：")
        self.label_deleteRow.grid(row=17, column=0)
        # Adding combobox drop down list

        self.button_deleteColumn = ttk.Button(self.frame3, text="削除", width=27, command=self.delete_row)
        self.button_deleteColumn.grid(row=17, column=1)

        # 行交換
        self.label_switchRow = ttk.Label(self.frame3, text="行交換：")
        self.label_switchRow.grid(row=18, column=0)
        # Adding combobox drop down list

        self.button_switchColumn = ttk.Button(self.frame3, text="交換", width=27, command=self.switch_row)
        self.button_switchColumn.grid(row=18, column=1)

        # Frame4
        """
        self.tree = ttk.Treeview(self.frame4, height=30)
        self.tree.grid(row=0, column=0)
        tree_h_scroll = ttk.Scrollbar(
            self.frame4,
            orient=tk.HORIZONTAL,
            command=self.tree.xview
        )

        tree_v_scroll = ttk.Scrollbar(
            self.frame4,
            orient=tk.VERTICAL,
            command=self.tree.yview
        )
        self.tree['xscrollcommand'] = tree_h_scroll.set
        self.tree['yscrollcommand'] = tree_v_scroll.set
        # self.tree.grid(row=0, column=0, sticky=tk.N + tk.S + tk.E + tk.W)
        tree_h_scroll.grid(row=1, column=0, sticky=tk.EW)
        tree_v_scroll.grid(row=0, column=1, sticky=tk.NS)
        """

    def inner_join_window(self):
        window = tk.Toplevel(self)
        window.title("INNER JOIN")  # ウィンドウタイトル
        window.geometry("800x380")  # ウィンドウサイズ(幅x高さ)
        frame_inner_join = ttk.Frame(window)
        frame_inner_join.grid(row=0, column=0, padx=20, pady=20)
        # スクロールバーの作成
        scrollbar_file_names = tk.Scrollbar(frame_inner_join, orient=tk.VERTICAL, command=tk.Listbox.yview)

        # リストボックス、スクロールバーの配置
        self.file_names_listbox2 = tk.Listbox(frame_inner_join, width=40, height=15, selectmode=MULTIPLE,
                                              yscrollcommand=scrollbar_file_names.set, background="darkgray")
        self.file_names_listbox2.grid(row=1, column=0)
        scrollbar_file_names.config(command=self.file_names_listbox2.yview)
        scrollbar_file_names.grid(row=1, column=1, sticky=[tk.N, tk.S])

        get_content = self.file_names_listbox.get(0, tkinter.END)
        for x in get_content:
            self.file_names_listbox2.insert(tk.END, x)

        # label
        label_choose = ttk.Label(frame_inner_join, text="結合する二つのファイルを選択して下さい")
        label_choose.grid(row=0, column=0)
        confirm_button = ttk.Button(frame_inner_join, text="決定", width=20, command=self.create_df_join)
        confirm_button.grid(row=2, column=0)

        # 間にspace1
        label_space1 = ttk.Label(frame_inner_join, text="           ")
        label_space1.grid(row=0, column=2)

        # 一つ目のデータフレームのカラムリストリストボックス
        scrollbar_df1 = tk.Scrollbar(frame_inner_join, orient=tk.VERTICAL, command=tk.Listbox.yview)

        # リストボックス、スクロールバーの配置
        self.df1_column_listbox = tk.Listbox(frame_inner_join, width=25, height=15, selectmode=tkinter.SINGLE,
                                             exportselection=0, yscrollcommand=scrollbar_df1.set)
        self.df1_column_listbox.grid(row=1, column=3)
        scrollbar_df1.config(command=self.df1_column_listbox.yview)
        scrollbar_df1.grid(row=1, column=4, sticky=[tk.N, tk.S])

        # 間にspace2
        label_space2 = ttk.Label(frame_inner_join, text="           ")
        label_space2.grid(row=0, column=5)

        # 二つ目のデータフレームのカラムリストリストボックス
        scrollbar_df2 = tk.Scrollbar(frame_inner_join, orient=tk.VERTICAL, command=tk.Listbox.yview)

        # リストボックス、スクロールバーの配置
        self.df2_column_listbox = tk.Listbox(frame_inner_join, width=25, height=15, selectmode=tkinter.SINGLE,
                                             exportselection=0,
                                             yscrollcommand=scrollbar_df2.set)
        self.df2_column_listbox.grid(row=1, column=6)
        scrollbar_df2.config(command=self.df2_column_listbox.yview)
        scrollbar_df2.grid(row=1, column=7, sticky=[tk.N, tk.S])

        # JOIN確定ボタン
        button_join_execute = ttk.Button(frame_inner_join, text="INNER JOIN", width=15, command=self.inner_join_execute)
        button_join_execute.grid(row=2, column=5)

    # inner join execute method
    def inner_join_execute(self):
        try:
            df1 = self.df_for_join[0]
            df2 = self.df_for_join[1]
            left_on = self.df2_column_listbox.get(self.df2_column_listbox.curselection())
            right_on = self.df1_column_listbox.get(self.df1_column_listbox.curselection())
            self.df = pd.merge(df1, df2, left_on=left_on, right_on=right_on, how="inner")
            self.create_table()
            self.reset_index()
            self.count_row_column(self.df)
            self.show_table(self.df)
            self.show_table_info()
            self.combobox_config()
        except:
            messagebox.showerror("エラー", "選択したコラムで結合はできません")

    # JOINように選ばれたファイルをdf化
    def create_df_join(self):
        try:
            self.df_for_join.clear()
            r = 0
            get_content = []
            for x in self.file_names_listbox2.curselection():
                get_content.append(self.file_names_listbox2.get(x))

            if len(get_content) != 2:
                messagebox.showinfo("エラー", "ファイルを二つ選択してください")
                return
            for i in get_content:
                if i.endswith(".xlsx"):
                    self.df_for_join.append(pd.read_excel(i))
                elif i.endswith(".csv"):
                    self.df_for_join.append(pd.read_csv(i))
                df_columns = list(self.df_for_join[-1].columns.values)
                j = 0
                for x in df_columns:
                    if r == 0:
                        if j == 0:
                            self.df1_column_listbox.delete(0, tk.END)
                            self.df1_column_listbox.insert(tk.END, x)
                            j += 1
                        else:
                            self.df1_column_listbox.insert(tk.END, x)
                    elif r == 1:
                        if j == 0:
                            self.df2_column_listbox.delete(0, tk.END)
                            self.df2_column_listbox.insert(tk.END, x)
                            j += 1
                        else:
                            self.df2_column_listbox.insert(tk.END, x)
                r += 1

        except EOFError:
            messagebox.showinfo("エラー", "ファイルの形式が間違っています")

    def add_window(self):
        if self.column_names:
            window = tk.Toplevel(self)
            window.title("行追加")  # ウィンドウタイトル
            window.geometry("800x380")  # ウィンドウサイズ(幅x高さ)

            frame = ttk.Frame(window)
            # bottom_frame = ttk.Frame(window)
            # frame.pack(fill=tkinter.BOTH, expand=1)
            frame.pack(fill=tkinter.BOTH, expand=1)
            # bottom_frame.pack(side='bottom')
            blank1 = ttk.Label(frame, text="       ")
            blank1.pack()
            # blank1.grid(row=0, column=0)
            label_explanation = ttk.Label(frame, text="値を入力して、追加ボタンをクリック")
            label_explanation.pack()
            # label_explanation.place(x=300, y=0)
            blank2 = ttk.Label(frame, text="       ")
            blank2.pack()
            # blank2.grid(row=1, column=0)
            canvas = tk.Canvas(frame)
            # canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.pack(side=tkinter.LEFT, expand=1, fill=tkinter.BOTH)
            h_scroll = ttk.Scrollbar(window, orient=tk.HORIZONTAL, command=canvas.xview)
            h_scroll.config(command=canvas.xview)
            canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
            canvas['xscrollcommand'] = h_scroll.set
            h_scroll.pack(side=tkinter.BOTTOM, fill=tkinter.X)

            second_frame = ttk.Frame(canvas)
            canvas.create_window((0, 0), window=second_frame, anchor="nw")

            # blank3.grid(row=4, column=0)
            button_add_row = ttk.Button(window, text="追加", width=15, command=self.add_row)
            button_add_row.pack(side="bottom")
            # button_add_row.grid(row=5, column=4)

            i = 1
            # blank4 = ttk.Label(canvas, text="       ")
            # blank4.grid(row=2, column=0)
            # use dictionary?
            for column in self.column_names:
                if column == "index":
                    pass
                else:
                    list_row = []
                    list_row.append(ttk.Label(second_frame, text=column))
                    self.list_entry.append(ttk.Entry(second_frame))
                    list_row[-1].grid(row=2, column=i)
                    self.list_entry[-1].grid(row=3, column=i)
                i = i + 1

            # scroll_x = ttk.Scrollbar(window, orient=tk.HORIZONTAL, command=window.xview)
            # self.tree.grid(row=0, column=0, sticky=tk.N + tk.S + tk.E + tk.W)
            # scroll_x.grid(row=6, column=0, sticky=tk.EW)

        else:
            messagebox.showinfo("エラー", "データが読み込まれていません")

    def add_row(self):
        # 最後のindexを入手し、＋１した値を追加する行のindexとして使用する
        new_index = self.rows_count
        new_elements = [new_index]
        for x in self.list_entry:
            new_elements.append(x.get())
        self.df.loc[len(self.df.index)] = new_elements
        self.count_row_column(self.df)
        self.show_table(self.df)
        print(self.df)

    # 　消した後のindexの変化をどうするか考える。
    def delete_row(self):
        try:
            selected_items = self.tree.selection()
            values = []
            if not selected_items:
                messagebox.showinfo("エラー", "削除する行が選択されていません")
                return
            for x in selected_items:
                values.append(self.tree.item(x)['values'])
            # 削除したい行のindexのリストを作成
            index_to_remove = []
            for x in values:
                index_to_remove.append(x[0])
            # 実際に削除する
            for x in index_to_remove:
                # self.df = self.df.drop([self.df.index[x]])
                self.df = self.df[self.df['index'] != x]
            self.reset_index()
            self.count_row_column(self.df)
            self.show_table(self.df)
        except:
            messagebox.showinfo("エラー", "ファイルが読み込まれていません")

    def switch_row(self):
        try:
            selected_items = self.tree.selection()
            values = []
            if not selected_items:
                messagebox.showinfo("エラー", "交換する行が選択されていません")
                return
            elif len(selected_items) != 2:
                messagebox.showinfo("エラー", "交換する行を二つ選択してください")
                return
            for x in selected_items:
                values.append(self.tree.item(x)['values'])
            # 削除したい行のindexのリストを作成
            index_to_switch = []
            for x in values:
                index_to_switch.append(x[0])
            a, b = self.df.iloc[index_to_switch[0]].copy(), self.df.iloc[index_to_switch[1]].copy()
            self.df.iloc[index_to_switch[0]], self.df.iloc[index_to_switch[1]] = b, a
            self.reset_index()
            self.show_table(self.df)
        except:
            messagebox.showinfo("エラー", "ファイルが読み込まれていません")

    def reset_index(self):
        try:
            del self.df["index"]
        except:
            pass
        index = list(range(len(self.df)))
        self.df.insert(loc=0, column='index', value=index)

    def select_file_dropbox(self):
        name = filedialog.askopenfilename()
        self.file_names_listbox.insert("end", name)

    def drop_inside_listbox(self, event):
        if event.data[-1] == "}":
            if event.data[0] == "{":
                event.data = event.data[:-1]
                event.data = event.data[1:]
            else:
                pass
        else:
            pass
        self.file_names_listbox.insert("end", event.data)

    def select_file(self, event):
        try:
            file_name = self.file_names_listbox.get(self.file_names_listbox.curselection())
            if file_name.endswith(".xlsx"):
                self.df = pd.read_excel(file_name)
            elif file_name.endswith(".csv"):
                self.df = pd.read_csv(file_name)
            messagebox.showinfo("確認", "Excelファイル読み込み完了")
        except:
            messagebox.showinfo("エラー", "ファイルの形式が間違っています")
        try:
            del self.tree
        except:
            pass
        index = list(range(len(self.df)))
        self.df.insert(loc=0, column='index', value=index)
        self.create_table()
        self.show_table_info()
        self.show_table(self.df)
        print(self.df)

    """
    def load_excel_file(self):
        try:

            self.df = pd.read_excel("./" + self.entry1.get() + ".xlsx", sheet_name=str(self.entry2.get()))
            # print(self.df)
            self.show_table_info()


        except:
            messagebox.showinfo("エラー", "入力されたファイル名かシート名が正しくありません")
    """

    # 表情報の表示
    def create_table(self):
        self.tree = ttk.Treeview(self.frame4, height=30)
        self.tree.grid(row=0, column=0)
        tree_h_scroll = ttk.Scrollbar(
            self.frame4,
            orient=tk.HORIZONTAL,
            command=self.tree.xview
        )

        tree_v_scroll = ttk.Scrollbar(
            self.frame4,
            orient=tk.VERTICAL,
            command=self.tree.yview
        )
        self.tree['xscrollcommand'] = tree_h_scroll.set
        self.tree['yscrollcommand'] = tree_v_scroll.set
        # self.tree.grid(row=0, column=0, sticky=tk.N + tk.S + tk.E + tk.W)
        tree_h_scroll.grid(row=1, column=0, sticky=tk.EW)
        tree_v_scroll.grid(row=0, column=1, sticky=tk.NS)

        self.button_show_df = ttk.Button(self.frame5, text="表全体を表示", width=25, command=self.show_current_df)
        self.button_show_df.grid(row=0, column=0)

        self.button_to_excel = ttk.Button(self.frame5, text="Excelファイル形式で保存", width=25, command=self.to_excel)
        self.button_to_excel.grid(row=0, column=1)

        self.button_to_csv = ttk.Button(self.frame5, text="CSVファイル形式で保存", width=25, command=self.to_csv)
        self.button_to_csv.grid(row=0, column=2)

    def to_excel(self):
        file = filedialog.asksaveasfile(mode='w', defaultextension=".xlsx")
        self.df.to_excel(file)

    def to_csv(self):
        file = filedialog.asksaveasfile(mode='w', defaultextension=".xlsx")
        self.df.to_csv(file)

    def show_current_df(self):
        self.show_table(self.df)
        self.show_table_info()

    def show_table_info(self):
        self.column_names = list(self.df.columns.values)
        # indexは表には常に表示されるが、削除できないようにカラム一覧からは消しておく
        self.column_names.pop(0)
        print(self.column_names)
        # 行数、列数の設定
        self.count_row_column(self.df)
        # comboboxに要素を追加
        self.combobox_config()

        # リストボックスにリストを追加
        # リストを空にする
        self.listbox.delete(0, tkinter.END)
        for x in self.column_names:
            self.listbox.insert(tk.END, x)

    def show_table(self, data):

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree["column"] = list(data.columns)

        self.tree.column("# 0", anchor=tkinter.CENTER, stretch=tkinter.NO, width=0)

        # self.tree.column("# 1", anchor=tkinter.CENTER, stretch=tkinter.NO, width=100)
        # self.tree.heading("# 1", text="ID")
        i = 1
        for column in self.tree["column"]:
            self.tree.heading("# " + str(i), text=column)
            self.tree.column("# " + str(i), anchor=tkinter.CENTER, stretch=tkinter.NO, width=100)
            i = i + 1
        rows = data.to_numpy().tolist()
        print(rows)
        for row in rows:
            self.tree.insert("", "end", values=row)
        self.count_row_column(data)

    """
    def load_excel_file(self):
        try:

            self.df = pd.read_excel("./" + self.entry1.get() + ".xlsx", sheet_name=str(self.entry2.get()))
            print(self.df)

            # 表情報の表示

            self.column_names = list(self.df.columns.values)
            print(self.column_names)
            # 行数、列数の設定
            self.count_row_column()
            # comboboxに要素を追加
            self.combobox_config()

            # リストボックスにリストを追加
            # リストを空にする
            self.listbox.delete(0, tkinter.END)
            for x in self.column_names:
                self.listbox.insert(tk.END, x)
            messagebox.showinfo("確認", "Excelファイル読み込み完了")
        except:
            messagebox.showinfo("エラー", "入力されたファイル名かシート名が正しくありません")
    """

    # 全選択する
    def select_all(self):
        self.listbox.select_set(0, tk.END)

    # 全クリアする
    def clear_all(self):
        self.listbox.select_clear(0, tk.END)

    # 確定する
    def confirm(self):
        self.selected.clear()
        try:
            for i in self.listbox.curselection():
                self.selected.append(self.listbox.get(i, i)[0])
                self.df_selected = self.df[self.selected].copy()
                self.df_selected.insert(0, 'index', self.df['index'])
                self.show_table(self.df_selected)
            if self.selected:
                messagebox.showinfo("選択完了", "選択したコラムでDataFrameを作成しました")
            else:
                messagebox.showinfo("エラー", "コラムを選択して下さい")
            print(self.df_selected)
        except:
            messagebox.showinfo("エラー", "ファイルが読み込まれていません")

    # 探索する
    def search(self):

        if self.selected:
            keyword = self.entry_keyWord.get()
            if keyword:
                self.df_CONCAT = self.df_selected[self.selected[0]].astype(str)
                for x in range(1, len(self.selected)):
                    self.df_CONCAT += self.df_selected[self.selected[x]].astype(str) + " "

                df_contain_value = self.df_selected[self.df_CONCAT.str.contains(keyword)]
                print(df_contain_value)
                self.show_table(df_contain_value)
                rows_count_contain_value = df_contain_value.shape[0]

                self.label_rowsKeywordResult.config(text=str(rows_count_contain_value) + " 行")

            else:
                messagebox.showinfo("エラー", "検索ワードが入力されていません")
        else:
            messagebox.showinfo("エラー", "カラムを選択して下さい")

    # 列追加  #既に存在するコラム名は追加できないようにする
    def add_column(self):
        column_new = self.entry_newColumn.get()
        try:
            if column_new:
                if column_new in self.df.columns:
                    messagebox.showinfo("エラー", "同じ名前のカラムが既に存在しています")
                else:
                    self.df[column_new] = " "
                    self.column_names.append(column_new)
                    self.listbox.insert(tk.END, column_new)
                    self.combobox_config()
                    self.count_row_column(self.df)
                    self.show_table(self.df)
            else:
                if self.df is not None:
                    messagebox.showinfo("エラー", "追加するカラム名を入力してください")
                else:
                    messagebox.showinfo("エラー", "ファイルが読み込まれていません")
        except:
            messagebox.showinfo("エラー", "ファイルが読み込まれていません")

    # 列削除
    def delete_column(self):
        try:
            column_delete = self.column_combobox1.get()
            if column_delete:
                if column_delete in self.column_names:
                    pass
                else:
                    messagebox.showinfo("エラー", "選択されたコラムは存在しません")
                    return
            else:
                messagebox.showinfo("エラー", "削除するコラムが選択されていません")
                return
            idx = self.column_names.index(column_delete)
            self.listbox.delete(idx)
            del self.df[column_delete]
            self.column_names.remove(column_delete)
            self.combobox_config()
            self.count_row_column(self.df)
            self.show_table(self.df)

        except:
            messagebox.showinfo("エラー", "ファイルが読み込まれていません")

    # 列交換
    def switch_column(self):
        col1 = self.column_combobox2.get()
        col2 = self.column_combobox3.get()

        if col1:
            if col2:
                pass
            else:
                messagebox.showinfo("エラー", "交換するコラムが選択されていません")
                return
        else:
            messagebox.showinfo("エラー", "交換するコラムが選択されていません")
            return

        self.df[col1], self.df[col2] = self.df[col2], self.df[col1]
        self.df.rename(columns={col1: col2, col2: col1}, inplace=True)
        x, y = self.column_names.index(col1), self.column_names.index(col2)
        self.column_names[y], self.column_names[x] = self.column_names[x], self.column_names[y]
        self.show_table(self.df)
        self.combobox_config()
        self.listbox.delete(0, tkinter.END)
        for x in self.column_names:
            self.listbox.insert(tk.END, x)

    # 行追加  #既に存在するコラム名は追加できないようにする
    # 行追加の場合、各列の値をどう入力するのか考える
    # 新しいwindowを開いて各列の情報を入力させる？

    # 行数列数設定
    def count_row_column(self, data):
        self.rows_count = data.shape[0]
        self.columns_count = data.shape[1]
        self.cell_count = self.rows_count * self.columns_count
        self.label_rowCount.config(text=str(self.rows_count) + " 行")
        self.label_columnCount.config(text=str(self.columns_count - 1) + " 列")
        self.label_cellCount.config(text=str(self.cell_count) + " セル")

    # combobox config
    def combobox_config(self):
        self.column_combobox1['values'] = self.column_names
        self.column_combobox2['values'] = self.column_names
        self.column_combobox3['values'] = self.column_names


def main():
    root = TkinterDnD.Tk()
    app = DataManagement(root)  # Inherit
    app.mainloop()


if __name__ == "__main__":
    main()
