from pathlib import Path
import sys, os, shutil
import xlwings as xw
import pandas as pd
import numpy as np
from PIL import Image
from typing import Optional, Tuple
import ttkbootstrap as ttk
import sqlite3
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import *
from ttkbootstrap.tableview import *
from ttkbootstrap.toast import ToastNotification
from tkinter.filedialog import askopenfilename, askopenfilenames
import tkinter as tk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

HEADER = [
    "Date",
    "Dept",
    "LU",
    "Technology",
    "Loss Name",
    "Evidence",
    "Component/ Part",
    "Chronology",
    "Countermeasure",
    "Object Part",
    "SAP#/OEM part number#: ",
    "Down time:",
    "Repair time:",
    "ACC  Accident or damage",
]


class Model:
    def __init__(self, db):
        # creating database connection
        self.con = sqlite3.connect(db)
        self.cur = self.con.cursor()

        # SQL queries to create tables
        bde = """
        CREATE TABLE IF NOT EXISTS bde (
            bde_id text PRIMARY KEY,
            date_bde date,
            dept text,
            link_up text,
            technology text,
            loss_name text,
            evidence text,
            component,
            chronology text,
            countermeasure text,
            object_part text,
            part_number text,
            down_time Integer,
            repair_time Integer,
            accident_damage text
        )
        """

        countermeasure = """
        CREATE TABLE IF NOT EXISTS countermeasure (
            cm_id text PRIMARY KEY,
            bde_id text,
            date_bde date,
            dept text,
            link_up text,
            technology text,
            loss_name text,
            evidence text,
            component,
            chronology text,
            countermeasure text,
            object_part text,
            part_number text,
            down_time Integer,
            repair_time Integer,
            accident_damage text,
            cm text,
            pic text,
            due_date text,
            status text
        )
        """

        # cursor executions
        self.cur.execute(bde)
        self.cur.execute(countermeasure)
        self.con.commit()

    def get_main_df(self):
        df = pd.read_sql_query(
            """SELECT * FROM countermeasure""", con=self.con
        ).sort_values(by=["cm_id"], ascending=False)

        df.loc[:, "date_bde"] = pd.to_datetime(df["date_bde"]).dt.date

        return df

    def get_main_df(self):
        df = pd.read_sql_query(
            """SELECT * FROM countermeasure""", con=self.con
        ).sort_values(by=["cm_id"], ascending=False)

        df.loc[:, "date_bde"] = pd.to_datetime(df["date_bde"]).dt.date

        return df


class Model_bde(Model):
    def __init__(self, db):
        super().__init__(db)

    def insert_BDE_record(self, args):
        query = self.__prepquery("INSERT INTO bde VALUES (%s)", range(len(args)))
        self.cur.execute(
            query,
            args,
        )
        self.con.commit()

    def delete_BDE_record(self, bde_id):
        query = "DELETE FROM bde WHERE bde_id = ?;"
        self.cur.execute(query, (bde_id,))
        self.con.commit()

    def __prepquery(self, qrystr, args):
        # replacement = "NULL, " + ", ".join("?" * len(args))
        replacement = ", ".join("?" * len(args))
        return qrystr % replacement

    def get_dataframe(self):
        self.df_bde = pd.read_sql_query(
            """SELECT * FROM bde""", con=self.con
        ).sort_values(by=["bde_id"], ascending=False)
        self.df_bde.loc[:, "date_bde"] = pd.to_datetime(
            self.df_bde["date_bde"], dayfirst=False
        ).dt.date
        df_BDE = self.df_bde[
            [
                "bde_id",
                "date_bde",
                "link_up",
                "technology",
                "loss_name",
                # "component",
                # "down_time",
                # "repair_time",
            ]
        ]
        return df_BDE

    def get_sub_id_count(self, date_bde):
        query = """SELECT * FROM bde """
        df = pd.read_sql_query(query, con=self.con)
        filtered_df = df[df["date_bde"] == str(date_bde)]
        return len(filtered_df["date_bde"])

    def get_data_header(self):
        df = self.get_dataframe()
        return df.columns.to_list()

    def get_data_value(self):
        df = self.get_dataframe()
        return df.values.tolist()

    def is_id_exists(self):
        ids = pd.read_sql_query("""SELECT bde_id FROM bde""", con=self.con)
        print(ids)

    def read_details_values(self, id):
        query = """SELECT * FROM bde"""
        df = pd.read_sql_query(query, con=self.con)
        details_values = df[df["bde_id"] == str(id)].values.tolist()[0]
        return details_values


class Model_cm(Model):
    def __init__(self, db):
        super().__init__(db)

    def insert_CM_record(self, args):
        query = self.__prepquery(
            "INSERT INTO countermeasure VALUES (%s)", range(len(args))
        )
        self.cur.execute(
            query,
            args,
        )
        self.con.commit()

    def __prepquery(self, qrystr, args):
        # replacement = "NULL, " + ", ".join("?" * len(args))
        replacement = ", ".join("?" * len(args))
        return qrystr % replacement

    def view(self):
        self.cur.execute("SELECT * FROM countermeasure")
        rows = self.cur.fetchall()
        return rows

    def update(self, status: str, cm_id: str):
        sql_insert_query = """UPDATE countermeasure SET status=? WHERE cm_id=?"""
        self.cur.execute(
            sql_insert_query,
            (
                status,
                cm_id,
            ),
        )
        self.con.commit()

    def delete(self, bde_id):
        self.cur.execute("DELETE FROM countermeasure WHERE bde_id=?", (bde_id,))
        self.con.commit()

    def get_dataframe(self):
        df = pd.read_sql_query(
            """SELECT * FROM countermeasure""", con=self.con
        ).sort_values(by=["cm_id"], ascending=False)

        df.loc[:, "date_bde"] = pd.to_datetime(df["date_bde"]).dt.date
        df_cm = df[
            [
                "cm_id",
                "bde_id",
                # "date_bde",
                # "link_up",
                # "technology",
                "loss_name",
                "cm",
                "pic",
                "due_date",
                "status",
            ]
        ]
        return df_cm

    def get_data_header(self):
        df = self.get_dataframe()
        return df.columns.to_list()

    def get_data_value(self):
        df = self.get_dataframe()
        return df.values.tolist()

    def get_cm_count(self):
        query = """SELECT * FROM countermeasure """
        df = pd.read_sql_query(query, con=self.con)
        return len(df)


class View(ttk.Window):
    app_icon = "bde_icon.ico"

    def __init__(self) -> None:
        super().__init__(themename="superhero")

        # widgets.
        self.title("BDE Tracker Cell 5")
        self.file_name = ttk.StringVar(value="")

        self.sidebar = SideBar(self, bootstyle=DARK)
        self.sidebar.pack(side=LEFT, expand=False, fill=Y)

        self.default_page = DefaultPage(self)
        self.default_page.pack(side=LEFT, expand=True, fill=BOTH)

        self.unpack_default_page_child()

        self.home_page = HomePage(self.default_page)
        self.home_page.pack(expand=True, fill=BOTH, anchor=N)
        self.table_bde = BDE_Page(self.default_page)
        self.table_cm = CM_Page(self.default_page)
        self.form_page = Form_Page(self.default_page)

    def unpack_default_page_child(self):
        for child in self.default_page.winfo_children():
            child.pack_forget()

    @property
    def add_sidebar_button(self) -> ttk.Button:
        return self.sidebar.home_button

    @property
    def bde_sidebar_button(self) -> ttk.Button:
        return self.sidebar.bde_button

    @property
    def cm_sidebar_button(self) -> ttk.Button:
        return self.sidebar.cm_button

    @property
    def extract_sidebar_button(self) -> ttk.Button:
        return self.sidebar.extract_button

    @property
    def sql_sidebar_button(self) -> ttk.Button:
        return self.sidebar.sql_button

    @property
    def bde_table(self):
        return self.table_bde

    @property
    def cm_table(self):
        return self.table_cm

    @property
    def form_input(self):
        return self.form_page

    @property
    def dashboard(self):
        return self.home_page

    @property
    def browse_button(self):
        return self.form_page.btn_browse

    @property
    def input_button(self):
        return self.form_page.btn_input

    @property
    def delete_button(self):
        return self.table_bde.button_delete

    @property
    def add_button(self):
        return self.table_bde.button_add


class SideBar(ttk.Frame):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.config(padding=5)

        self.sidebar_image = load_image_tk("bde_image.jpg", (100, 100))

        # Logo Image
        self.logo_image = ttk.Label(self)
        self.logo_image.config(image=self.sidebar_image)
        self.logo_image.pack(side=ttk.TOP, padx=10, pady=(40, 20))

        # Add Data Button
        self.home_button = ttk.Button(self)
        self.home_button.config(text="Dashboard")
        self.home_button.config(cursor="hand2")
        self.home_button.config(width=14)
        self.home_button.config(bootstyle=PRIMARY)
        self.home_button.pack(side=ttk.TOP, padx=10, pady=10)

        # BDE Button
        self.bde_button = ttk.Button(self)
        self.bde_button.config(text="BDE")
        self.bde_button.config(cursor="hand2")
        self.bde_button.config(width=14)
        self.bde_button.config(bootstyle=PRIMARY)
        self.bde_button.pack(side=ttk.TOP, padx=10, pady=10)

        # Countermeasure Button
        self.cm_button = ttk.Button(self)
        self.cm_button.config(text="Countermeasure")
        self.cm_button.config(cursor="hand2")
        self.cm_button.config(width=14)
        self.cm_button.config(bootstyle=PRIMARY)
        self.cm_button.pack(side=ttk.TOP, padx=10, pady=10)

        # Extract Button
        self.extract_button = ttk.Button(self)
        self.extract_button.config(text="Extract")
        self.extract_button.config(cursor="hand2")
        self.extract_button.config(width=14)
        self.extract_button.config(bootstyle=PRIMARY)
        # self.extract_button.pack(side=ttk.TOP, padx=10, pady=10)

        # SQL Button
        self.sql_button = ttk.Button(self)
        self.sql_button.config(text="SQL")
        self.sql_button.config(cursor="hand2")
        self.sql_button.config(width=14)
        self.sql_button.config(bootstyle="primary-outline")
        # self.sql_button.pack(side=ttk.TOP, padx=10, pady=10)


class DefaultPage(ttk.Frame):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.config(padding=5)

        # # header.
        # self.header_container = ttk.Frame(self)
        # self.header_container.pack(side=ttk.TOP, fill=ttk.X)

        # body.
        self.body_container = ttk.Frame(self)
        self.body_container.pack(
            side=ttk.TOP, fill=ttk.BOTH, expand=ttk.YES, anchor="n"
        )

        # # footer.
        # self.footer_container = ttk.Frame(self)
        # self.footer_container.pack(side=ttk.TOP, fill=ttk.X)

    def destroy_body_content_child(self):
        for child in self.body_container.winfo_children():
            print(child)
            child.destroy()


class HomePage(DefaultPage):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.dashbord_frame = ttk.Frame(self.body_container)
        self.dashbord_frame.pack(
            side=TOP, padx=10, pady=2, fill=BOTH, expand=YES, anchor=N
        )

        self.create_axis()

    def create_axis(self):

        fig = plt.figure()
        fig.subplots_adjust(
            left=0.05, right=0.95, top=0.9, bottom=0.15, wspace=0.3, hspace=0.4
        )
        fig.set_figwidth(17)
        fig.set_figheight(9)
        fig.set_dpi(70)
        self.ax1 = fig.add_subplot(231)
        self.ax2 = fig.add_subplot(232)
        self.ax3 = fig.add_subplot(233)
        self.ax4 = fig.add_subplot(223)
        self.ax5 = fig.add_subplot(224)

        canvas1 = FigureCanvasTkAgg(fig, self.dashbord_frame).get_tk_widget()
        canvas1.pack(fill=BOTH, expand=YES)

    def destroy_axis(self):
        self.ax1.cla()
        self.ax2.cla()
        self.ax3.cla()
        self.ax4.cla()
        self.ax5.cla()

    def show_plot(self, df: pd.DataFrame):
        # subplot 1
        rect1 = self.ax1.bar(
            df["status"].value_counts(ascending=False).to_dict().keys(),
            df["status"].value_counts(ascending=False).to_dict().values(),
        )
        self.ax1.set_title("CM Status")
        self.ax1.margins(0.05, 0.1)

        for rect in rect1:
            height = rect.get_height()
            x_pos = rect.get_x() + rect.get_width() / 2
            y_pos = height
            self.ax1.annotate(
                text=str(height),
                xy=(x_pos, y_pos),
                ha="center",
                xytext=(0, 3),
                textcoords="offset points",
            )

        # subplot 2
        df_bde = df.drop_duplicates(["bde_id"])
        df_bde.loc[:, "date_bde"] = (
            df_bde["date_bde"]
            .apply(lambda x: str(x.year) + "-" + str(x.month).zfill(2))
            .to_frame()
        )

        rect2 = self.ax2.bar(
            df_bde["link_up"].value_counts(ascending=False).to_dict().keys(),
            df_bde["link_up"].value_counts(ascending=False).to_dict().values(),
        )
        self.ax2.set_title("Link Up")
        self.ax2.margins(0.05, 0.1)

        for rect in rect2:
            height = rect.get_height()
            x_pos = rect.get_x() + rect.get_width() / 2
            y_pos = height
            self.ax2.annotate(
                text=str(height),
                xy=(x_pos, y_pos),
                ha="center",
                xytext=(0, 3),
                textcoords="offset points",
            )

        # subplot 4
        df_grouped = df_bde.groupby("date_bde")["link_up"].count().reset_index()
        rect4 = self.ax4.bar(
            df_grouped["date_bde"].values.tolist(),
            df_grouped["link_up"].values.tolist(),
        )
        self.ax4.set_title("History")
        self.ax4.tick_params(axis="x", labelrotation=90)
        self.ax4.margins(0.05, 0.1)

        for rect in rect4:
            height = rect.get_height()
            x_pos = rect.get_x() + rect.get_width() / 2
            y_pos = height
            self.ax4.annotate(
                text=str(height),
                xy=(x_pos, y_pos),
                ha="center",
                xytext=(0, 3),
                textcoords="offset points",
            )

        # subplot 5
        rect5 = self.ax5.bar(
            df_bde["technology"].value_counts(ascending=False).to_dict().keys(),
            df_bde["technology"].value_counts(ascending=False).to_dict().values(),
        )
        self.ax5.set_title("Technology")
        self.ax5.tick_params(axis="x", labelrotation=90)
        self.ax5.margins(0.05, 0.1)

        for rect in rect5:
            height = rect.get_height()
            x_pos = rect.get_x() + rect.get_width() / 2
            y_pos = height
            self.ax5.annotate(
                text=str(height),
                xy=(x_pos, y_pos),
                ha="center",
                xytext=(0, 3),
                textcoords="offset points",
            )


class BDE_Page(DefaultPage):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.table_BDE = Tableview(
            master=self.body_container,
            paginated=False,
            searchable=True,
            bootstyle=DARK,
            autofit=True,
            height=35,
        )
        self.table_BDE.pack(side=LEFT, fill=BOTH, expand=YES, padx=0, pady=0)

        self.details = ttk.Frame(master=self.body_container)
        self.details.pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)

        self.frame_form = ScrolledFrame(self.details, height=750, width=600, padding=10)
        self.frame_form.pack(side=TOP, fill=BOTH, expand=YES)

        self.button_frame = ttk.Frame(master=self.details)
        self.button_frame.pack(side=BOTTOM, fill=X, expand=YES, padx=10, pady=10)

        self.button_add = ttk.Button(
            master=self.button_frame, text="ADD", style=SUCCESS
        )
        self.button_add.pack(side=LEFT, fill=X, expand=YES, padx=10, pady=10)

        self.button_delete = ttk.Button(
            master=self.button_frame, text="DELETE", style=DANGER
        )
        self.button_delete.pack(side=LEFT, fill=X, expand=YES, padx=10, pady=10)

        self.__craete_details()

    def __craete_details(self):
        header = HEADER
        for i in range(len(header)):
            self.book_id_frame = ttk.Labelframe(self.frame_form, height=50)
            self.book_id_frame.config(text=header[i])
            self.book_id_frame.pack(
                side=ttk.TOP, fill=ttk.X, expand=True, padx=10, pady=1
            )

    def set_values(self, coldata, rowdata):
        self.table_BDE.build_table_data(coldata=coldata, rowdata=rowdata)
        # self.table_BDE.build_table_data(coldata=coldata, rowdata=rowdata)

    def binding(self, sequence: str, func):
        self.table_BDE.view.bind(sequence, func)

    def get_row_data(self, selected):
        selected = self.table_BDE.view.selection()

        records = []
        for iid in selected:
            record: TableRow = self.table_BDE.iidmap.get(iid)
            records.append(record.values)

        return records[0]

    def destroy_content_child(self):
        for child in self.body_container.winfo_children():
            child.destroy()

    def get_values(self):
        values = []
        for children in self.frame_form.winfo_children():
            for child in children.winfo_children():
                value = self.__get_values(child)
                values.append(value)
        return values

    def set_detail_values(self, value: list):
        # self.reset_detail_values()
        # list_field = self.frame_form.winfo_children()
        # for i in range(len(list_field)):
        #     self.__insert(list_field[i], value[i])

        self.reset_detail_values()
        list_field = self.frame_form.winfo_children()
        for i in range(len(list_field)):
            if i == 7 or i == 8:
                self.__insert_scrolled(list_field[i], value[i])
            else:
                self.__insert_entry(list_field[i], value[i])

    def reset_detail_values(self):
        list_field = self.frame_form.winfo_children()
        for i in range(len(list_field)):
            self.__clear(list_field[i])

    @staticmethod
    def __insert_entry(master: tk.Misc, value: str) -> None:
        label = ttk.Entry(master)
        label.insert(0, value)
        label.pack(side=ttk.TOP, fill=ttk.X, expand=True, padx=2, pady=2)

    @staticmethod
    def __insert_scrolled(master: tk.Misc, value: str) -> None:
        label = ScrolledText(master, width=100, height=5, state="normal")
        label.insert(END, value)
        label.pack(side=ttk.TOP, fill=ttk.X, expand=True, padx=2, pady=2)

    @staticmethod
    def __clear(master: tk.Misc) -> None:
        for children in master.winfo_children():
            children.destroy()

    @staticmethod
    def __get_values(master: ttk.Entry):
        return master.get()


class CM_Page(DefaultPage):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.table_CM = Tableview(
            master=self.body_container,
            paginated=False,
            searchable=True,
            bootstyle=DARK,
            autofit=True,
            height=35,
            autoalign=True,
        )

        self.table_CM.pack(side=LEFT, fill=BOTH, expand=YES, padx=0, pady=0)

    def set_values(self, coldata, rowdata):
        self.table_CM.build_table_data(coldata=coldata, rowdata=rowdata)

    def binding(self, sequence: str, func):
        self.table_CM.view.bind(sequence, func)

    def get_row_data(self, selected):
        selected = self.table_CM.view.selection()

        records = []
        for iid in selected:
            record: TableRow = self.table_CM.iidmap.get(iid)
            records.append(record.values)

        return records[0]


class Form_Page(DefaultPage):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)

        self.file_name = ttk.StringVar()
        self.header = HEADER
        self.data = ["", "", "", "", "", "", "", "", "", "", "", "", "", ""]

        browse_frame = ttk.Frame(self.body_container)
        browse_frame.pack(side=TOP, padx=10, pady=2, fill=X, expand=False)
        self.btn_browse = ttk.Button(
            master=browse_frame,
            text="Browse File",
            # command=self.browse_btn_click,
            bootstyle=PRIMARY,
            width=13,
        )
        self.btn_browse.pack(side=LEFT, padx=10, pady=10)

        self.form_input = ttk.Entry(master=browse_frame, textvariable=self.file_name)
        self.form_input.pack(side=LEFT, padx=5, fill=X, expand=YES)

        self.create_field(self.data)

        self.btn_input = ttk.Button(
            master=self.body_container,
            text="Input",
            # command=self.input_btn_click,
            bootstyle=SUCCESS,
            width=13,
        )
        self.btn_input.pack(side=TOP, padx=10, pady=(20, 10))

    def create_field(self, list_value):
        self.frame_form = ScrolledFrame(
            self.body_container, height=700, width=600, padding=10
        )
        self.frame_form.pack(side=TOP, fill=BOTH, expand=YES)

        for i in range(len(self.header)):
            self.book_id_frame = ttk.Labelframe(self.frame_form, height=50)
            self.book_id_frame.config(text=self.header[i])
            self.book_id_frame.pack(
                side=ttk.TOP, fill=ttk.X, expand=True, padx=10, pady=1
            )
        # self.create_entry()
        self.set_values(self.data)

    @staticmethod
    def __insert_entry(master: tk.Misc, value: str) -> None:
        label = ttk.Entry(master)
        label.insert(0, value)
        label.pack(side=ttk.TOP, fill=ttk.X, expand=True, padx=2, pady=2)

    @staticmethod
    def __insert_scrolled(master: tk.Misc, value: str) -> None:
        label = ScrolledText(master, width=100, height=5, state="normal")
        label.insert(END, value)
        label.pack(side=ttk.TOP, fill=ttk.X, expand=True, padx=2, pady=2)

    @staticmethod
    def __clear(master: tk.Misc) -> None:
        for children in master.winfo_children():
            children.destroy()

    @staticmethod
    def __get_values(master):
        if isinstance(master, ttk.Entry):
            return master.get()
        else:
            return master.get("1.0", tk.END)

    def get_values(self):
        values = []
        for children in self.frame_form.winfo_children():
            for child in children.winfo_children():
                value = self.__get_values(child)
                values.append(value)
        return values

    def set_values(self, value: list):
        self.reset_values()
        list_field = self.frame_form.winfo_children()
        for i in range(len(list_field)):
            if i == 7 or i == 8:
                self.__insert_scrolled(list_field[i], value[i])
            else:
                self.__insert_entry(list_field[i], value[i])

    def reset_values(self):
        list_field = self.frame_form.winfo_children()
        for i in range(len(list_field)):
            self.__clear(list_field[i])


class Controller:
    def __init__(self, view: View, model: list[Model]) -> None:
        self.model_bde: Model_bde = model[0]
        self.model_cm: Model_cm = model[1]
        self.view = view

        self.__bind_books_page()

        df_bde = self.model_cm.get_main_df()
        self.dashboard = self.view.dashboard
        self.dashboard.destroy_axis()
        self.dashboard.show_plot(df_bde)

    def __bind_books_page(self) -> None:
        dashboard_button = self.view.add_sidebar_button
        dashboard_button.config(command=self.dashboard_button_click)

        bde_button = self.view.bde_sidebar_button
        bde_button.config(command=self.bde_button_click)

        cm_button = self.view.cm_sidebar_button
        cm_button.config(command=self.cm_button_click)

        browse_button = self.view.browse_button
        browse_button.config(command=self.browse_button_click)

        input_button = self.view.input_button
        input_button.config(command=self.input_button_click)

        sql_button = self.view.sql_sidebar_button
        sql_button.config(command=self.sql_button_click)

        add_button = self.view.add_button
        add_button.config(command=self.add_button_click)

        delete_button = self.view.delete_button
        delete_button.config(command=self.delete_button_click)

    def dashboard_button_click(self):
        self.view.unpack_default_page_child()
        df_bde = self.model_cm.get_main_df()

        self.dashboard = self.view.dashboard
        self.dashboard.destroy_axis()
        self.dashboard.show_plot(df_bde)
        self.dashboard.pack(expand=True, fill="both")

    def add_button_click(self):
        self.view.unpack_default_page_child()

        self.form_input = self.view.form_input
        self.form_input.pack(expand=True, fill="y")

    def bde_button_click(self):
        head = self.model_bde.get_data_header()
        val = self.model_bde.get_data_value()

        self.view.unpack_default_page_child()

        self.bde_table = self.view.bde_table
        self.bde_table.set_values(coldata=head, rowdata=val)

        self.bde_table.binding("<<TreeviewSelect>>", self.bde_table_click)
        self.bde_table.pack(expand=True, fill="both")

    def cm_button_click(self):
        head = self.model_cm.get_data_header()
        val = self.model_cm.get_data_value()

        self.view.unpack_default_page_child()

        self.cm_table = self.view.cm_table
        self.cm_table.set_values(coldata=head, rowdata=val)
        self.cm_table.binding("<Double-1>", self.cm_table_double_click)
        self.cm_table.pack(expand=True, fill="both")

    def delete_button_click(self):
        selected_row = self.bde_table.table_BDE.view.selection()
        row_data = self.bde_table.get_row_data(selected_row)
        bde_id = row_data[0]
        self.model_bde.delete_BDE_record(bde_id)
        self.model_cm.delete(bde_id)
        self.bde_button_click()
        self.show_toast("Status has been successfully deleted", "success")

    def bde_table_click(self, event):
        selected_row = self.bde_table.table_BDE.view.selection()
        row_data = self.bde_table.get_row_data(selected_row)

        bde_id = row_data[0]
        detail_value = self.model_bde.read_details_values(bde_id)[1:]

        for i in range(len(detail_value)):
            if detail_value[i] == None or detail_value[i] == np.nan:
                detail_value[i] = ""

        # print(detail_value)
        self.bde_table.set_detail_values(detail_value)

        # print(detail_value)

    def cm_table_double_click(self, event):
        selected_row = self.cm_table.table_CM.view.selection()
        row_data = self.cm_table.get_row_data(selected_row)

        cm_id = row_data[0]
        if row_data[6] == "Open":
            status = "Close"
        else:
            status = "Open"

        self.model_cm.update(status=status, cm_id=cm_id)

        self.cm_button_click()
        self.show_toast("Status has been successfully updated", "success")

    def browse_button_click(self):
        try:
            self.filenames = askopenfilename(
                title="Select file", filetypes=(("Excel Files", "*.xls*"),)
            )

            self.view.form_page.file_name.set(self.filenames)

            self.data = get_extraction_data(self.filenames)
            for i in range(len(self.data)):
                if self.data[i] is None:
                    self.data[i] = ""

            self.view.form_input.set_values(self.data)

        except Exception as e:
            print(e)

    def input_button_click(self):
        field_values = self.view.form_input.get_values()
        try:

            if "" in field_values or field_values == []:
                self.show_toast(
                    "Data is incomplete, please complete the PM Card", "warning"
                )
            else:
                bde_values, cm_values = get_dataframe_values(
                    self.data, self.model_bde, self.model_cm
                )

                for data_bde in bde_values:
                    self.model_bde.insert_BDE_record(data_bde)

                for data_cm in cm_values:
                    self.model_cm.insert_CM_record(data_cm)

                self.view.form_input.reset_values()
                self.view.form_page.file_name.set("")

                copy_file(self.filenames)
                self.bde_button_click()
                self.show_toast("Data has been successfully entered", "success")

        except Exception as e:
            self.bde_button_click()
            self.show_toast(e, "warning")

    def sql_button_click(self):
        filenames = askopenfilenames(
            title="Select file", filetypes=(("Excel Files", "*.xls*"),)
        )
        for filename in filenames:
            df_bde = pd.read_excel(filename, sheet_name="Sheet1")
            for i in range(len(df_bde)):
                df_i = df_bde.loc[i, :].to_frame().transpose()
                df_i.loc[:, "Date"] = pd.to_datetime(df_i["Date"]).dt.date
                bde_values = df_i.values.tolist()
                for data_bde in bde_values:
                    self.model_bde.insert_BDE_record(data_bde)

            df_cm = pd.read_excel(filename, sheet_name="Sheet2")
            for i in range(len(df_cm)):
                df_i = df_cm.loc[i, :].to_frame().transpose()
                df_i.loc[:, "Date"] = pd.to_datetime(df_i["Date"]).dt.date
                cm_values = df_i.values.tolist()
                for data_cm in cm_values:
                    self.model_cm.insert_CM_record(data_cm)

    def show_toast(self, message, bootstyle):
        toast = ToastNotification(
            title="BDE App",
            message=message,
            duration=3000,
            bootstyle=bootstyle,
        )
        toast.show_toast()


def main():
    create_folder("Database")
    create_folder("Document BDE")
    db_path = Path("./Database/Database_BDE.db")
    model_bde = Model_bde(db_path)
    model_cm = Model_cm(db_path)
    model = [model_bde, model_cm]
    application = View()
    controller = Controller(application, model)
    application.mainloop()


def load_image_tk(
    path: Path, geometry: Optional[Tuple[int, int]] = None
) -> ttk.ImageTk.PhotoImage:
    image = Image.open(resource_path(path))

    if geometry is not None:
        return ttk.ImageTk.PhotoImage(image.resize(geometry))
    return ttk.ImageTk.PhotoImage(image)


def resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and for PyInstaller"""

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def get_script_folder() -> str:
    """get script folder, .py scropt or .exe file location"""

    if getattr(sys, "frozen", False):
        script_path = os.path.dirname(sys.executable)
    else:
        script_path = os.path.dirname(os.path.abspath(sys.modules["__main__"].__file__))
    return script_path


def create_folder(folder_name: str) -> None:
    folder_path = os.path.join(get_script_folder(), folder_name)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)


def copy_file(src):
    file_name = os.path.basename(src)
    dst = Path(f"./Document BDE/{file_name}")
    shutil.copyfile(src, dst)


def get_dataframe(data):
    header = HEADER
    df_bde = pd.DataFrame(data=data).transpose()
    df_bde.columns = header
    df_bde.loc[:, "Date"] = pd.to_datetime(df_bde["Date"]).dt.date

    return df_bde


def get_dataframe_values(data, model_bde: Model_bde, model_cm: Model_cm):
    # BDE
    df_bde = get_dataframe(data)
    print(df_bde.to_dict())
    lu = str(df_bde.loc[0, "LU"]).replace("-ID01", "")
    date = str(df_bde.loc[0, "Date"]).replace("-", "")
    no = model_bde.get_sub_id_count(df_bde.loc[0, "Date"]) + 1
    num = "{:02d}".format(no)

    temp_id = f"BDE_{lu}_{date}_{num}"

    df_bde["bde_id"] = temp_id

    df_bde.set_index(df_bde.columns[-1], inplace=True)
    df_bde.reset_index(inplace=True)

    bde_values = df_bde.values.tolist()

    # CUNTERMEASURE
    df_cm = df_bde
    df_cm.loc[:, "Countermeasure"] = df_cm["Countermeasure"].str.split("\n")
    df_cm = df_cm.explode(["Countermeasure"])
    df_cm = df_cm.reset_index(drop=True)
    df_cm[["Countermeasures", "PIC", "Due Date"]] = df_cm["Countermeasure"].str.split(
        "|", expand=True
    )
    df_cm["Countermeasures"] = df_cm["Countermeasures"].str.strip()
    df_cm["PIC"] = df_cm["PIC"].str.strip()
    df_cm["Due Date"] = df_cm["Due Date"].str.strip()
    df_cm["Status"] = "Open"

    no = model_cm.get_cm_count() + 1

    temp_id = []
    for i in range(len(df_cm)):
        temp = no + i
        num = "{:05d}".format(temp)
        x = f"CM_{num}"
        temp_id.append(x)

    df_cm["cm_id"] = temp_id

    df_cm.set_index(df_cm.columns[-1], inplace=True)
    df_cm.reset_index(inplace=True)

    cm_values = df_cm.values.tolist()

    return bde_values, cm_values


def get_extraction_data(filename: str) -> list:
    app = xw.App(visible=False)
    book = xw.Book(filename)
    sheet = book.sheets("Extraction")
    data = sheet.range("C4:P4").value
    book.close()
    app.quit()

    return data


if __name__ == "__main__":
    main()
