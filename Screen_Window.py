from tkinter import *
from tkinter import ttk
from tkmacosx import Button
import sqlite3

from reportlab.pdfgen import canvas
import webbrowser
import os
import pandas as pd


window = Tk()


class Reports:
    def generate_pdf(self):
        file = 'Client_Report.pdf'
        path_to_browser = webbrowser.get('open -a /Applications/Safari.app %s')
        path_to_browser.open_new_tab(file)

    def report_generator(self):
        self.canva = canvas.Canvas('Client_Report.pdf')

        self.codeRel = self.code_entry.get()
        self.first_nameRel = self.first_name_entry.get()
        self.last_nameRel = self.last_name_entry.get()
        self.idRel = self.id_entry.get()
        self.SSNRel = self.SSN_entry.get()
        self.PassportRel = self.Passport_entry.get()
        self.emailRel = self.email_entry.get()
        self.phoneRel = self.phone_entry.get()
        self.addressRel = self.address_entry.get()
        self.cityRel = self.city_entry.get()
        self.StateRel = self.State_entry.get()
        self.Zip_CodeRel = self.Zip_Code_entry.get()
        self.CountryRel = self.Country_entry.get()

        self.canva.setFont('Helvetica-Bold', 24)
        self.canva.drawString(200, 790, 'Client Information')
        self.canva.rect(20, 740, 550, 100, fill=False, stroke=True)

        self.canva.setFont('Helvetica-Bold', 18)
        self.canva.drawString(50, 700, 'Code: ')
        self.canva.drawString(50, 673, 'First Name: ')
        self.canva.drawString(50, 646, 'Last Name: ')
        self.canva.drawString(50, 619, 'ID: ')
        self.canva.drawString(50, 592, 'SSN: ')
        self.canva.drawString(50, 565, 'Passport: ')
        self.canva.drawString(50, 538, 'Email: ')
        self.canva.drawString(50, 511, 'Phone: ')
        self.canva.drawString(50, 484, 'Address: ')
        self.canva.drawString(50, 457, 'City: ')
        self.canva.drawString(50, 430, 'State: ')
        self.canva.drawString(50, 403, 'Zip Code: ')
        self.canva.drawString(50, 376, 'Country: ')

        self.canva.setFont('Helvetica-Bold', 18)
        self.canva.drawString(170, 700, self.codeRel)
        self.canva.drawString(170, 673, self.first_nameRel)
        self.canva.drawString(170, 646, self.last_nameRel)
        self.canva.drawString(170, 619, self.idRel)
        self.canva.drawString(170, 592, self.SSNRel)
        self.canva.drawString(170, 565, self.PassportRel)
        self.canva.drawString(170, 538, self.emailRel)
        self.canva.drawString(170, 511, self.phoneRel)
        self.canva.drawString(170, 484, self.addressRel)
        self.canva.drawString(170, 457, self.cityRel)
        self.canva.drawString(170, 430, self.StateRel)
        self.canva.drawString(170, 403, self.Zip_CodeRel)
        self.canva.drawString(170, 376, self.CountryRel)

        self.canva.showPage()
        self.canva.save()
        self.generate_pdf()


class Funcs:

    def clear_screen(self):
        """
        The function will be called by the clear widget which will clear all the 
        typed information showing on the input fields.
        :return: Blank entry fields.
        """
        self.code_entry.delete(0, END)
        self.first_name_entry.delete(0, END)
        self.last_name_entry.delete(0, END)
        self.id_entry.delete(0, END)
        self.SSN_entry.delete(0, END)
        self.Passport_entry.delete(0, END)
        self.email_entry.delete(0, END)
        self.phone_entry.delete(0, END)
        self.address_entry.delete(0, END)
        self.city_entry.delete(0, END)
        self.State_entry.delete(0, END)
        self.Zip_Code_entry.delete(0, END)
        self.Country_entry.delete(0, END)

    def db_connection(self):
        self.conn = sqlite3.connect('customers.db')
        self.cursor = self.conn.cursor()
        if self.cursor:
            print('Database connection established.')

    def db_disconnect(self):
        self.conn.close()
        print('Database connection closed.')

    def table_creation(self):
        self.db_connection()
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS VCF_Customers (
                Cod INTEGER PRIMARY KEY,
                First_Name CHAR(40) NOT NULL,
                Last_Name CHAR(40) NOT NULL,
                ID INTEGER(20),
                SSN INTEGER(20),
                Passport CHAR(20),
                Email CHAR(60),
                Phone INTEGER(20),
                Address CHAR(60),
                City CHAR(40),
                State CHAR(40),
                Zip_Code INTEGER(15),
                Country CHAR(40)
            );
        """)
        self.conn.commit()
        print('Database successfully created.')
        self.db_disconnect()

    def variables(self):
        self.code = self.code_entry.get()
        self.first_name = self.first_name_entry.get()
        self.last_name = self.last_name_entry.get()
        self.id = self.id_entry.get()
        self.SSN = self.SSN_entry.get()
        self.Passport = self.Passport_entry.get()
        self.email = self.email_entry.get()
        self.phone = self.phone_entry.get()
        self.address = self.address_entry.get()
        self.city = self.city_entry.get()
        self.State = self.State_entry.get()
        self.Zip_Code = self.Zip_Code_entry.get()
        self.Country = self.Country_entry.get()

    def add_client(self):
        self.variables()
        self.db_connection()
        self.cursor.execute(""" INSERT INTO VCF_Customers (first_name, last_name, id, SSN, Passport, email, phone, 
            address, city, State, Zip_Code, Country)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) """, (self.first_name, self.last_name, self.id, self.SSN,
                                                              self.Passport, self.email, self.phone, self.address,
                                                              self.city, self.State, self.Zip_Code,
                                                              self.Country))
        self.conn.commit()
        self.db_disconnect()
        self.select_list()
        self.clear_screen()

    def select_list(self):
        self.clients_table.delete(*self.clients_table.get_children())
        self.db_connection()
        table = self.cursor.execute(''' SELECT cod, first_name, last_name, id, SSN, Passport, email, phone, address,
            city, State, Zip_Code, Country FROM VCF_Customers
            ORDER BY first_name ASC; ''')

        for item in table:
            self.clients_table.insert('', END, value=item)
        self.db_disconnect()

    def search_client(self):
        self.db_connection()
        self.clients_table.delete(*self.clients_table.get_children())

        self.last_name_entry.insert(END, '%')
        last = self.last_name_entry.get()

        self.cursor.execute(
            '''SELECT cod, first_name, last_name, id, SSN, Passport, email, phone, address, city, State, Zip_Code,
            Country FROM VCF_Customers
            WHERE last_name LIKE '%s' ORDER BY last_name ASC''' % last)

        search_client = self.cursor.fetchall()
        for item in search_client:
            self.clients_table.insert('', END, values=item)

        self.clear_screen()
        self.db_disconnect()

    def on_double_click(self, event):
        self.clear_screen()
        self.clients_table.selection()

        for key in self.clients_table.selection():
            col1, col2, col3, col4, col5, col6, col7, \
            col8, col9, col10, col11, col12, col13 = self.clients_table.item(key, 'values')

            self.code_entry.insert(END, col1)
            self.first_name_entry.insert(END, col2)
            self.last_name_entry.insert(END, col3)
            self.id_entry.insert(END, col4)
            self.SSN_entry.insert(END, col5)
            self.Passport_entry.insert(END, col6)
            self.email_entry.insert(END, col7)
            self.phone_entry.insert(END, col8)
            self.address_entry.insert(END, col9)
            self.city_entry.insert(END, col10)
            self.State_entry.insert(END, col11)
            self.Zip_Code_entry.insert(END, col12)
            self.Country_entry.insert(END, col13)

    def delete_client(self):
        self.variables()
        self.db_connection()
        self.cursor.execute(''' DELETE FROM VCF_Customers WHERE cod = ? ''', (self.code))
        self.conn.commit()
        self.db_disconnect()
        self.clear_screen()
        self.select_list()

    def update_client(self):
        self.variables()
        self.db_connection()
        self.cursor.execute(''' UPDATE VCF_Customers SET first_name = ?, last_name = ?, id = ?, SSN = ?, Passport = ?,
            email = ?, phone = ?, address = ?, city = ?, State = ?, Zip_code = ?, Country = ? 
            WHERE cod = ? ''', (self.first_name, self.last_name, self.id, self.SSN, self.Passport, self.email,
                                self.phone, self.address, self.city, self.State, self.Zip_Code, self.Country,
                                self.code))
        self.conn.commit()
        self.db_disconnect()
        self.select_list()
        self.clear_screen()

    def generate_xml(self):
        os.system("open -a '/Applications/Microsoft Excel.app' 'VCF_Customers.xlsx'")

    def xml_report(self):
        self.clients_table.delete(*self.clients_table.get_children())
        self.db_connection()
        self.cursor.execute(''' SELECT cod, first_name, last_name, id, SSN, Passport, email, phone, address,
                    city, State, Zip_Code, Country FROM VCF_Customers
                    ORDER BY cod ASC ''')
        clients_table = self.cursor.fetchall()
        for item in clients_table:
            self.clients_table.insert('', END, values=item)
            clients_table = pd.DataFrame(clients_table, columns=['Code', 'First Name', 'Last Name', 'Id / RG',
                                                                 'CPF / SSN', 'Passport', 'E-mail', 'Phone',
                                                                 'Address', 'City', 'State', 'Zip Code', 'Country'])
            clients_table.to_excel('VCF_Customers.xlsx')
        self.db_disconnect()


class Application(Funcs, Reports):
    window: Tk

    def __init__(self) -> None:
        self.window = window
        self.screen()
        self.screen_frames()
        self.widgets_frame_1()
        self.table_frame_2()
        self.table_creation()
        self.select_list()
        self.Menus()
        self.xml_report()
        window.mainloop()

    def screen(self) -> str:
        """
        Defines the window characteristics for the application.
        Such as Title, Geometry, Size restrictions, Background color etc.
        :return: The customized window of the application with the specified parameters.
        """

        self.window.title("Customer Registration")
        self.window.configure(background='#1e3743')
        self.window.geometry('680x580')
        self.window.resizable(True, True)
        self.window.maxsize(None, None)
        self.window.minsize(width=600, height=550)

    def screen_frames(self):
        """
        Adds two separated frames into the window. Using relx, rely etc. so the parameters will
        adapt and actively follow the window as it's size changes by the user. 
        :return: Two customized frames inside the window with the specified parameters.
        """
        self.frame_1 = Frame(self.window, bd=4, bg='#dfe3ee', highlightbackground='#759fe6', highlightthickness=3)
        self.frame_1.place(relx=0.014, rely=0.02, relwidth=0.97, relheight=0.47)

        self.frame_2 = Frame(self.window, bd=4, bg='#dfe3ee', highlightbackground='#759fe6', highlightthickness=3)
        self.frame_2.place(relx=0.014, rely=0.5, relwidth=0.97, relheight=0.47)

    def widgets_frame_1(self) -> object:
        """
        Creating buttons for the Application.
        :return: Buttons that will to facilitate the application usage and navigation.
        """

        # Frame for Buttons on left side of the screen.
        self.canvas_bt = Canvas(self.frame_1, bd=0, bg='#1e3743', highlightbackground='gray',
                                highlightthickness=3)
        self.canvas_bt.place(relx=0.194497, rely=0.0015, relwidth=0.21497, relheight=0.19)

        # Clear
        self.bt_clear = Button(self.frame_1, text='Clear', bd=2, bg='#187db2', fg='white',
                               activebackground='#108ecb', activeforeground='white',
                               font=('verdana', 13, 'bold'), command=self.clear_screen)
        self.bt_clear.place(relx=0.2, rely=0.02, relwidth=0.1, relheight=0.15)

        # Search
        self.bt_search = Button(self.frame_1, text='Search', bd=2, bg='#187db2', fg='white',
                                activebackground='#108ecb', activeforeground='white',
                                font=('verdana', 13, 'bold'), command=self.search_client)
        self.bt_search.place(relx=0.3016, rely=0.02, relwidth=0.1, relheight=0.15)

        # Frame for Buttons on the right side of the screen.
        self.canvas_bt = Canvas(self.frame_1, bd=0, bg='#1e3743', highlightbackground='gray',
                                highlightthickness=3)
        self.canvas_bt.place(relx=0.594497, rely=0.0015, relwidth=0.31497, relheight=0.19)

        # New Customer Button
        self.bt_new = Button(self.frame_1, text='New', bd=2, bg='#187db2', fg='white',
                             activebackground='#108ecb', activeforeground='white',
                             font=('verdana', 13, 'bold'), command=self.add_client)
        self.bt_new.place(relx=0.6, rely=0.02, relwidth=0.1, relheight=0.15)

        # Update Button
        self.bt_update = Button(self.frame_1, text='Update', bd=2, bg='#187db2', fg='white',
                                activebackground='#108ecb', activeforeground='white',
                                font=('verdana', 13, 'bold'), command=self.update_client)
        self.bt_update.place(relx=0.7016, rely=0.02, relwidth=0.1, relheight=0.15)

        # Delete
        self.bt_delete = Button(self.frame_1, text='Delete', bd=2, bg='#187db2', fg='white',
                                activebackground='#108ecb', activeforeground='white',
                                font=('verdana', 13, 'bold'), command=self.delete_client)
        self.bt_delete.place(relx=0.8024495, rely=0.02, relwidth=0.1, relheight=0.15)

        # Creating Code Label & Input Box
        self.lb_code = Label(self.frame_1, text='Code', bg='#dfe3ee', fg='black')
        self.lb_code.place(relx=0.05, rely=0.05)

        self.code_entry = Entry(self.frame_1, bg='#187db2')
        self.code_entry.place(relx=0.05, rely=0.14, relwidth=0.06)

        # First Name Label & Input Box
        self.lb_first_name = Label(self.frame_1, text='First Name', bg='#dfe3ee', fg='black')
        self.lb_first_name.place(relx=0.05, rely=0.3)

        self.first_name_entry = Entry(self.frame_1, bg='#187db2')
        self.first_name_entry.place(relx=0.05, rely=0.4, relwidth=0.15)

        # Last Name Label & Entry Box
        self.lb_last_name = Label(self.frame_1, text='Last Name', bg='#dfe3ee', fg='black')
        self.lb_last_name.place(relx=0.22, rely=0.3)

        self.last_name_entry = Entry(self.frame_1, bg='#187db2')
        self.last_name_entry.place(relx=0.22, rely=0.4, relwidth=0.15)

        # Id Label & Input Box
        self.lb_id = Label(self.frame_1, text='Id', bg='#dfe3ee', fg='black')
        self.lb_id.place(relx=0.38, rely=0.3)

        self.id_entry = Entry(self.frame_1, bg='#187db2')
        self.id_entry.place(relx=0.38, rely=0.4, relwidth=0.13)

        # SSN Label & Input Box
        self.lb_ssn = Label(self.frame_1, text='SSN', bg='#dfe3ee', fg='black')
        self.lb_ssn.place(relx=0.52, rely=0.3)

        self.SSN_entry = Entry(self.frame_1, bg='#187db2')
        self.SSN_entry.place(relx=0.52, rely=0.4, relwidth=0.14)

        # Passport Label & Input Box
        self.lb_Passport = Label(self.frame_1, text='Passport', bg='#dfe3ee', fg='black')
        self.lb_Passport.place(relx=0.67, rely=0.3)

        self.Passport_entry = Entry(self.frame_1, bg='#187db2')
        self.Passport_entry.place(relx=0.67, rely=0.4, relwidth=0.11)

        # Email Label & Entry Box
        self.lb_email = Label(self.frame_1, text='Email', bg='#dfe3ee', fg='black')
        self.lb_email.place(relx=0.05, rely=0.52)

        self.email_entry = Entry(self.frame_1, bg='#187db2')
        self.email_entry.place(relx=0.05, rely=0.62, relwidth=0.4)

        # Phone Label & Input Box
        self.lb_phone = Label(self.frame_1, text='Phone', bg='#dfe3ee', fg='black')
        self.lb_phone.place(relx=0.5, rely=0.52)

        self.phone_entry = Entry(self.frame_1, bg='#187db2')
        self.phone_entry.place(relx=0.5, rely=0.62, relwidth=0.4)

        # Address Label & Input Box
        self.lb_address = Label(self.frame_1, text='Address', bg='#dfe3ee', fg='black')
        self.lb_address.place(relx=0.05, rely=0.74)

        self.address_entry = Entry(self.frame_1, bg='#187db2')
        self.address_entry.place(relx=0.05, rely=0.84, relwidth=0.35)

        # City Label & Input Box
        self.lb_city = Label(self.frame_1, text='City', bg='#dfe3ee', fg='black')
        self.lb_city.place(relx=0.41, rely=0.74)

        self.city_entry = Entry(self.frame_1, bg='#187db2')
        self.city_entry.place(relx=0.41, rely=0.84, relwidth=0.13)

        # State Label & Input Box
        self.lb_State = Label(self.frame_1, text='State', bg='#dfe3ee', fg='black')
        self.lb_State.place(relx=0.55, rely=0.74)

        self.State_entry = Entry(self.frame_1, bg='#187db2')
        self.State_entry.place(relx=0.55, rely=0.84, relwidth=0.05)

        # Zip Code Label & Input Box
        self.lb_Zip_Code = Label(self.frame_1, text='Zip Code', bg='#dfe3ee', fg='black')
        self.lb_Zip_Code.place(relx=0.62, rely=0.74)

        self.Zip_Code_entry = Entry(self.frame_1, bg='#187db2')
        self.Zip_Code_entry.place(relx=0.62, rely=0.84, relwidth=0.11)

        # Country Label & Input Box
        self.lb_Country = Label(self.frame_1, text='Country', bg='#dfe3ee', fg='black')
        self.lb_Country.place(relx=0.74, rely=0.74)

        self.Country_entry = Entry(self.frame_1, bg='#187db2')
        self.Country_entry.place(relx=0.74, rely=0.84, relwidth=0.15)

    def table_frame_2(self):
        self.clients_table = ttk.Treeview(self.frame_2, height=3, columns=('col1', 'col2', 'col3', 'col4', 'col5',
                                                                           'col6', 'col7', 'col8', 'col9', 'col10',
                                                                           'col11', 'col12', 'col13'))
        self.clients_table.heading('#0', text='')
        self.clients_table.heading('#1', text='Code')
        self.clients_table.heading('#2', text='First Name')
        self.clients_table.heading('#3', text='Last Name')
        self.clients_table.heading('#4', text='ID')
        self.clients_table.heading('#5', text='SSN')
        self.clients_table.heading('#6', text='Passport')
        self.clients_table.heading('#7', text='Email')
        self.clients_table.heading('#8', text='Phone')
        self.clients_table.heading('#9', text='Address')
        self.clients_table.heading('#10', text='City')
        self.clients_table.heading('#11', text='State')
        self.clients_table.heading('#12', text='Zip Code')
        self.clients_table.heading('#13', text='Country')

        self.clients_table.column('#0', width=1)
        self.clients_table.column('#1', width=2)
        self.clients_table.column('#2', width=30)
        self.clients_table.column('#3', width=30)
        self.clients_table.column('#4', width=27)
        self.clients_table.column('#5', width=13)
        self.clients_table.column('#6', width=27)
        self.clients_table.column('#7', width=50)
        self.clients_table.column('#8', width=15)
        self.clients_table.column('#9', width=50)
        self.clients_table.column('#10', width=27)
        self.clients_table.column('#11', width=20)
        self.clients_table.column('#12', width=37)
        self.clients_table.column('#13', width=15)

        self.clients_table.place(relx=0.01, rely=0.01, relwidth=0.95, relheight=0.85)

        self.scroll_table = Scrollbar(self.frame_2, orient='vertical')
        self.clients_table.configure(yscroll=self.scroll_table.set)
        self.scroll_table.place(relx=0.97, rely=0.01, relwidth=0.02, relheight=0.85)

        self.clients_table.bind('<Double-1>', self.on_double_click)

    def Menus(self):
        menubar = Menu(self.window)
        self.window.config(menu=menubar)
        File = Menu(menubar)
        Export = Menu(menubar)

        def quit(): self.window.destroy()

        menubar.add_cascade(label='Options', menu=File)
        menubar.add_cascade(label='Reports', menu=Export)

        File.add_command(label='New', command=self.add_client)
        File.add_command(label='Update', command=self.update_client)
        File.add_command(label='Search', command=self.search_client)
        File.add_command(label='Delete', command=self.delete_client)
        File.add_command(label='Quit', command=quit)

        Export.add_command(label='Export to PDF', command=self.report_generator)
        Export.add_command(label='Export to Excel', command=self.generate_xml)


if __name__ == '__main__':
    Application()
