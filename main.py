from PyQt5 import QtCore,QtGui,QtWidgets,uic
from PyQt5.QtGui import QIcon,QPixmap
from PyQt5.uic import loadUiType
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import sys
import sqlconn as sqc
from sqlalchemy import asc
import win32com.client
import xlwings.constants
import datetime
from PyQt5.uic import loadUiType
import os
import shutil
import inflect
ui,_=loadUiType('stock_inventory_system.ui')
procurement_ui, _ = loadUiType('procurement.ui')
accounts_ui, _ = loadUiType('accounts.ui')
item_ui, _ = loadUiType('item.ui')
department_ui, _ = loadUiType('department.ui')
global settings_account_table
global settings_department_table
global price_list_table
global gen_tab
def find_cell(scsht, tofind, srow=1, scol=1, lrow=1, lcol=1, search_order='by_row', search_sheet='advance',
              lookin='formulas'):
    """
    :param scsht: source sheet object #cannot be None
    :param tofind: string #cannot be None
    :param srow: start row, integer #default 1
    :param scol: start column, integer #default 1
    :param lrow: last row, integer, #default 1
    :param lcol: last column, integer #default 1
    :param search_order: 'by_row' or 'by_col', string #default 'by_row'
    :param search_sheet: 'advance' or 'basic' if basic no lrow or lcol needed. #default is advance
    :param lookin: 'formulas','values' or 'comments' #default formulas
    :return: dict['row'] and dict['col']
    """
    search_order_dict = {'by_row': xlwings.constants.SearchOrder.xlByRows,
                         'by_col': xlwings.constants.SearchOrder.xlByColumns}
    lookin_dict = {'values': xlwings.constants.FindLookIn.xlValues,
                   'formulas': xlwings.constants.FindLookIn.xlFormulas,
                   'comments': xlwings.constants.FindLookIn.xlComments}
    if search_sheet == 'advance':
        cell = scsht.Range(scsht.Cells(srow, scol), scsht.Cells(lrow, lcol)).Find(What=tofind.strip(),
                                                                                  LookAt=xlwings.constants.LookAt.xlWhole,
                                                                                  LookIn=lookin_dict[lookin],
                                                                                  SearchOrder=xlwings.constants.SearchOrder.xlByRows,
                                                                                  MatchCase=False)
        if cell is None:
            cell = scsht.Range(scsht.Cells(srow, scol), scsht.Cells(lrow, lcol)).Find(What=tofind.strip(),
                                                                                      LookAt=xlwings.constants.LookAt.xlPart,
                                                                                      LookIn=lookin_dict[lookin],
                                                                                      SearchOrder=search_order_dict[
                                                                                          search_order],
                                                                                      MatchCase=False)
        print('{} at , row = {}, col = {}'.format(tofind, cell.Row, cell.Column))
    elif search_sheet == 'basic':
        cell = scsht.Cells.Find(What="*", After=scsht.Cells(1, 1), SearchOrder=search_order_dict[search_order],
                                SearchDirection=xlwings.constants.SearchDirection.xlPrevious)
    cellpos = {'row': cell.Row, 'col': cell.Column}
    return cellpos

def remove_non_digits(input):
    output = ''.join(c for c in input if c.isdigit())
    return output

def remove_digits(input):
    output = ''.join(c for c in input if c.isdigit() == False)
    return output

def is_number(s):
    try:
        float(s)
        return True
    except:
        return False


class Item_Dialogue(QDialog,item_ui):
    edit_id = 0
    operationType = ''

    def __init__(self,parent=None):
        super(Item_Dialogue,self).__init__(parent)
        self.setupUi(self)

    def ShowDialogue(self,id,item,price,operationType=''):
        self.item.setText(item)
        self.price.setText(price)
        self.edit_id = id
        self.operationType = operationType
        self.buttonBox.accepted.connect(self.ok_button)

    def ok_button(self):
        engine = sqc.Database().engine
        pgso_price_list = sqc.Database().pgso_price_list
        conn = engine.connect()

        if self.operationType == 'edit':
            self.item_label.setText('Edit Account')
            s = pgso_price_list.update().where(pgso_price_list.c.id == self.edit_id).\
                values(item = self.item.text(),
                       price = self.price.text())
            conn.execute(s)
            self.show_price_list()

        elif self.operationType == 'add':
            self.item_label.setText('Add Account')
            s = pgso_price_list.insert().values(
                item=self.item.text(),
                price=self.price.text())
            conn.execute(s)
            self.show_price_list()
        conn.close()

    def show_price_list(self):
        global price_list_table
        price_list_table.setRowCount(0)
        engine = sqc.Database().engine
        pgso_price_list = sqc.Database().pgso_price_list
        conn= engine.connect()
        #admin_table
        s = pgso_price_list.select().order_by(asc(pgso_price_list.c.item))
        s_value = conn.execute(s)
        table = price_list_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))
            table.resizeColumnsToContents()
        conn.close()

class Accounts_Dialogue(QDialog,accounts_ui):
    edit_id = 0
    operationType = ''

    def __init__(self,parent=None):
        super(Accounts_Dialogue,self).__init__(parent)
        self.setupUi(self)

    def ShowDialogue(self,id,username,password,operationType=''):
        self.username.setText(username)
        self.password.setText(password)
        self.edit_id = id
        self.operationType = operationType
        self.buttonBox.accepted.connect(self.ok_button)

    def ok_button(self):
        engine = sqc.Database().engine
        pgso_admin = sqc.Database().pgso_admin
        conn = engine.connect()

        if self.operationType == 'edit':
            self.account_label.setText('Edit Account')
            s = pgso_admin.update().where(pgso_admin.c.userid == self.edit_id).\
                values(username = self.username.text(),
                       password = self.password.text())
            conn.execute(s)
            self.show_settings()

        elif self.operationType == 'add':
            self.account_label.setText('Add Account')
            s = pgso_admin.insert().values(
                username=self.username.text(),
                password=self.password.text())
            conn.execute(s)
            self.show_settings()
        conn.close()

    def show_settings(self):
        global settings_account_table
        settings_account_table.setRowCount(0)
        engine = sqc.Database().engine
        pgso_admin = sqc.Database().pgso_admin
        conn= engine.connect()
        #admin_table
        s = pgso_admin.select().order_by(asc(pgso_admin.c.username))
        s_value = conn.execute(s)
        table = settings_account_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))
        conn.close()

class Department_Dialogue(QDialog,department_ui):
    edit_id = 0
    operationType = ''

    def __init__(self,parent=None):
        super(Department_Dialogue,self).__init__(parent)
        self.setupUi(self)

    def ShowDialogue(self,id,type,name,operationType=''):
        index = self.department_type.findText(type)
        if index >= 0:
            self.department_type.setCurrentIndex(index)

        self.department_name.setText(name)
        self.edit_id = id
        self.operationType = operationType
        self.buttonBox.accepted.connect(self.ok_button)

    def ok_button(self):
        engine = sqc.Database().engine
        pgso_department = sqc.Database().pgso_department
        conn = engine.connect()

        if self.operationType == 'edit':
            self.department_label.setText('Edit Department')
            s = pgso_department.update().where(pgso_department.c.id == self.edit_id).\
                values(type = self.department_type.currentText(),
                       name = self.department_name.text())
            conn.execute(s)
            self.show_settings()

        elif self.operationType == 'add':
            self.department_label.setText('Add Department')
            s = pgso_department.insert().values(
                type=self.department_type.currentText(),
                name=self.department_name.text())
            conn.execute(s)
            self.show_settings()
        conn.close()

    def show_settings(self):
        global settings_department_table
        settings_department_table.setRowCount(0)
        engine = sqc.Database().engine
        pgso_department = sqc.Database().pgso_department
        conn= engine.connect()
        #admin_table
        s = pgso_department.select().order_by(asc(pgso_department.c.type))
        s_value = conn.execute(s)
        table = settings_department_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))
        conn.close()

class Procurement_Dialogue(QDialog,procurement_ui):

    def __init__(self,parent=None):
        super(Procurement_Dialogue,self).__init__(parent)
        self.setupUi(self)
        self.Handle_UI()
        self.Handle_Buttons()

    def Handle_UI(self):
        self.show_department_name()

    def Handle_Buttons(self):
        self.department_type.currentTextChanged.connect(self.show_department_name)
        self.open_button.clicked.connect(self.open_button_action)
        self.ok_button.clicked.connect(self.ok_button_action)
        self.cancel_button.clicked.connect(lambda : self.close())

    def show_department_name(self):
        self.department_name.clear()
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_department = sqc.Database().pgso_department
        s = pgso_department.select().where(pgso_department.c.type == self.department_type.currentText())
        s_value = conn.execute(s)
        for val in s_value:
            self.department_name.addItem(val[2])

    def open_button_action(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        stpfile, _ = QFileDialog.getOpenFileName(self, "Open STP File",
                                                   "",
                                                   "XL Files (*.xlsx);;All Files (*)", options=options)
        self.attached_file.setText(stpfile)

    def ok_button_action(self):
        try:
            excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
            scbk = excel.Workbooks.Open(r'{}'.format(self.attached_file.text()), ReadOnly=True)
            scsht = scbk.Worksheets[1]
            lrow = find_cell(scsht, '', search_order='by_row', search_sheet='basic')['row']
            lcol = find_cell(scsht, '', search_order='by_col', search_sheet='basic')['col']
            start_index = find_cell(scsht,'GENERAL DESCRIPTION',lrow=lrow,lcol=lcol)
            end_row = find_cell(scsht,'NOTE: Technical Specifications for each Item',lrow=lrow,lcol=lcol)['row']
            quantity_col = find_cell(scsht,'QUANTITY',lrow=lrow,lcol=lcol)['col']

            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_department = sqc.Database().pgso_department
            s = pgso_department.select().where(pgso_department.c.type == self.department_type.currentText()).\
                where(pgso_department.c.name == self.department_name.currentText())
            s_value = conn.execute(s)
            id = 0
            for val in s_value:
                id = val[0]

            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_procurement = sqc.Database().pgso_procurement
            s = pgso_procurement.insert().values(
                department_id = id,
                date_archived = datetime.datetime.utcnow(),
                status = 'pr'
            )
            conn.execute(s)
            po_id = 0
            s = pgso_procurement.select()
            s_value = conn.execute(s)
            for val in s_value :
                po_id = val[0]
            conn.close()

            for row in range(start_index['row'] + 1,end_row):
                col = start_index['col']
                if scsht.Cells(row,col).Value is not None:
                    description = str(scsht.Cells(row,col).Value)
                    quantity = 0
                    unit = ''
                    try:
                        quantity = int(remove_non_digits(str(scsht.Cells(row,quantity_col).Value)))
                        unit = remove_digits(str(scsht.Cells(row,quantity_col).Value)).strip()
                    except:
                        pass



                    engine2 = sqc.Database().engine
                    conn2 = engine2.connect()
                    price_list_dict = {}
                    pgso_price_list = sqc.Database().pgso_price_list
                    s= pgso_price_list.select()
                    s_value = conn2.execute(s)

                    for val in s_value:
                        if is_number(val[2]):
                            price_list_dict.update({val[1].strip().lower() : float(val[2])})

                    try:
                        unit_cost = price_list_dict[description.strip().lower()]
                    except:
                        unit_cost = 0

                    pgso_procurement_data = sqc.Database().pgso_procurement_data
                    s = pgso_procurement_data.insert().values(
                        description = description,
                        quantity = quantity,
                        unit = unit,
                        unit_cost = unit_cost,
                        po_id = po_id,
                    )
                    conn2.execute(s)

            scbk.Close(SaveChanges=False)
            excel.Quit()

            shutil.copyfile(os.path.abspath(self.attached_file.text()), os.path.abspath(os.getcwd() + '\\excel\\' + str(po_id) + '.xlsx'))

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Inserted")
            msg.setInformativeText('Data Inserted to the Database')
            msg.setWindowTitle("PGSO Purchase Request")
            msg.exec_()
            self.close()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Something Went Wrong , Do it Again')
            msg.setWindowTitle("Error")
            msg.exec_()
            self.close()

class MainApp(QMainWindow,ui):

    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.Handle_UI_Changes()
        self.Handle_Globals()
        self.Handle_Buttons()

    def Handle_Globals(self):
        global settings_account_table
        settings_account_table = self.settings_account_table
        global settings_department_table
        settings_department_table = self.settings_department_table
        global price_list_table
        price_list_table = self.price_list_table

    def Handle_UI_Changes(self):
        self.menu_widget.setVisible(False)
        self.tabWidget.tabBar().setVisible(False)
        self.tabWidget.setCurrentIndex(0)
        self.home_logo.setEnabled(False)
        ##menu
        self.menu_logo.setEnabled(False)
        ##settings
        self.settings_account_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.settings_account_table.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.settings_account_table.setColumnHidden(0,True)
        self.settings_department_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.settings_department_table.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.settings_department_table.setColumnHidden(0,True)
        self.price_list_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.price_list_table.setSelectionBehavior(QtWidgets.QTableView.SelectRows)
        self.price_list_table.setColumnHidden(0,True)

        ##pr
        self.show_pr_name()
        self.show_po_name()
        ##genpr
        self.gen_table_widget.setColumnHidden(0, True)

    def Handle_Buttons(self):
        ##login
        self.login_button.clicked.connect(self.login_button_action)
        self.login_username.textChanged.connect(lambda: self.login_error_message.setText(''))
        self.login_password.textChanged.connect(lambda: self.login_error_message.setText(''))
        ##menu
        self.menu_logout.clicked.connect(self.menu_logout_action)
        self.menu_transaction.clicked.connect(self.menu_transaction_action)
        self.menu_settings.clicked.connect(self.menu_settings_action)
        #transactionm
        self.transaction_procurement.clicked.connect(self.transaction_procurement_action)
        self.transaction_purchase_request.clicked.connect(self.transaction_purchase_request_action)
        self.transaction_purchase_order.clicked.connect(self.transaction_purchase_order_action)
        ##purchase_request
        self.pr_generate.clicked.connect(self.pr_generate_action)
        self.pr_type.currentTextChanged.connect(self.show_pr_name)
        self.pr_type.currentTextChanged.connect(self.show_pr_list)
        self.pr_name.currentTextChanged.connect(self.show_pr_list)
        self.pr_open_in_excel.clicked.connect(self.pr_open_in_excel_action)
        self.pr_delete.clicked.connect(self.pr_delete_action)
        ##genpr
        self.gen_cancel_button.clicked.connect(self.gen_cancel_button_action)
        self.gen_purchase_order.clicked.connect(self.gen_purchase_order_action)
        #purchase_order
        self.po_type.currentTextChanged.connect(self.show_po_name)
        self.po_type.currentTextChanged.connect(self.show_po_list)
        self.po_name.currentTextChanged.connect(self.show_po_list)
        self.po_edit.clicked.connect(self.po_edit_action)
        self.po_generate_pr.clicked.connect(self.po_generate_pr_action)
        self.po_generate_po.clicked.connect(self.po_generate_po_action)
        self.po_delete.clicked.connect(self.po_delete_action)
        ##settings
        self.settings_add_account.clicked.connect(self.settings_add_account_action)
        self.settings_edit_account.clicked.connect(lambda : self.settings_edit_account_action(self.settings_account_table))
        self.settings_delete_account.clicked.connect(lambda : self.settings_delete_account_action(self.settings_account_table))
        self.settings_add_department.clicked.connect(self.settings_add_department_action)
        self.settings_edit_department.clicked.connect(lambda : self.settings_edit_department_action(self.settings_department_table))
        self.settings_delete_department.clicked.connect(lambda : self.settings_delete_department_action(self.settings_department_table))
        ##price_list
        self.menu_price_list.clicked.connect(self.menu_price_list_action)
        self.price_list_add.clicked.connect(self.price_list_add_action)
        self.price_list_edit.clicked.connect(lambda : self.price_list_edit_action(self.price_list_table))
        self.price_list_delete.clicked.connect(lambda : self.price_list_delete_action(self.price_list_table))

    def refresh_application(self):
        self.menu_widget.setVisible(False)
        self.tabWidget.setCurrentIndex(0)
        self.login_username.setText('')
        self.login_password.setText('')

    ##login
    def login_button_action(self):
        username = self.login_username.text()
        password = self.login_password.text()

        engine = sqc.Database().engine
        pgso_admin = sqc.Database().pgso_admin
        conn = engine.connect()
        s = pgso_admin.select()
        s_value = conn.execute(s)

        for val in s_value:
            if str(username).lower() == str(val[1]).lower() and str(password).lower() == str(val[2]).lower():
                self.tabWidget.setCurrentIndex(1)
                self.menu_widget.setVisible(True)
            else:
                self.login_username.setText('')
                self.login_password.setText('')
                self.login_error_message.setText('Wrong username or password!!')

        conn.close()

    ##menu
    def menu_logout_action(self):
        self.refresh_application()

    def menu_transaction_action(self):
        self.tabWidget.setCurrentIndex(2)

    def menu_settings_action(self):
        self.tabWidget.setCurrentIndex(6)
        self.show_settings()

    ##transaction
    def transaction_procurement_action(self):
        d = Procurement_Dialogue(self)
        d.show()

    def transaction_purchase_request_action(self):
        self.tabWidget.setCurrentIndex(3)
        self.show_pr_list()

    def transaction_purchase_order_action(self):
        self.tabWidget.setCurrentIndex(5)
        self.show_po_list()

    ##pr
    def show_pr_name(self):
        self.pr_name.clear()
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_department = sqc.Database().pgso_department
        s = pgso_department.select().where(pgso_department.c.type == self.pr_type.currentText())
        s_value = conn.execute(s)
        for val in s_value:
            self.pr_name.addItem(val[2])

    pr_dict = {}
    def show_pr_list(self):
        self.pr_list.clear()
        self.pr_dict = {}
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_department = sqc.Database().pgso_department
        s = pgso_department.select().where(pgso_department.c.type == self.pr_type.currentText()). \
            where(pgso_department.c.name == self.pr_name.currentText())
        s_value = conn.execute(s)
        id = 0
        for val in s_value:
            id = val[0]

        pgso_procurement = sqc.Database().pgso_procurement
        s = pgso_procurement.select().where(pgso_procurement.c.department_id == id).where(pgso_procurement.c.status == 'pr')
        s_value = conn.execute(s)

        for val in s_value:
            self.pr_dict.update({'{}_(#{})_{}'.format(self.pr_name.currentText(),val[0],val[2]) : val[0]})

        for key in self.pr_dict.keys():
            self.pr_list.addItem(key)

    def pr_generate_action(self):
        try:
            self.tabWidget.setCurrentIndex(4)
            self.show_gen_table_widget()
        except:
            pass

    def show_gen_table_widget(self):
        global gen_tab
        gen_tab = 3
        temp = self.pr_list.currentItem().text()
        temp2 = temp.split('(')[0] +temp.split(')')[1]
        self.gen_title.setText(temp2)
        po_id = self.pr_dict[self.pr_list.currentItem().text()]
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_procurement_data = sqc.Database().pgso_procurement_data
        s = pgso_procurement_data.select().where(pgso_procurement_data.c.po_id == po_id).order_by(pgso_procurement_data.c.description)
        s_value = conn.execute(s)
        table = self.gen_table_widget
        table.setRowCount(0)
        for val in s_value :
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[3])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 3, QTableWidgetItem(str(val[2])))
            table.setItem(row_position, 4, QTableWidgetItem(str(val[4])))
            table.resizeColumnsToContents()
            table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)

    def pr_open_in_excel_action(self):
        try:
            os.startfile(os.getcwd() + '\\excel\\' + str(self.pr_dict[self.pr_list.currentItem().text()]) + '.xlsx')
        except:
            pass

    def pr_delete_action(self):
        try:
            po_id = self.pr_dict[self.pr_list.currentItem().text()]
            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_procurement = sqc.Database().pgso_procurement
            s = pgso_procurement.delete().where(pgso_procurement.c.id == po_id)
            conn.execute(s)
            pgso_procurement_data = sqc.Database().pgso_procurement_data
            s = pgso_procurement_data.delete().where(pgso_procurement_data.c.po_id == po_id)
            conn.execute(s)
            os.remove(os.getcwd() + '\\excel\\' + str(self.pr_dict[self.pr_list.currentItem().text()]) + '.xlsx')
            self.show_pr_list()
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Success")
            msg.setInformativeText('File Deleted Properly')
            msg.setWindowTitle("Information")
            msg.exec_()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Something went wrong in file deletion')
            msg.setWindowTitle("Error")
            msg.exec_()

    def gen_purchase_order_action(self):
        try:
            table = self.gen_table_widget
            end_row = table.rowCount()
            if gen_tab == 3:
                po_id = self.pr_dict[self.pr_list.currentItem().text()]
            elif gen_tab == 5:
                po_id = self.po_dict[self.po_list.currentItem().text()]
            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_procurement = sqc.Database().pgso_procurement
            pgso_procurement_data = sqc.Database().pgso_procurement_data
            for i in range(0,end_row):
                id = table.item(i,0).text()
                unit = table.item(i,1).text()
                description = table.item(i,2).text()
                try:
                    quantity = int(table.item(i,3).text())
                except:
                    quantity = 0
                try:
                    unit_cost = int(table.item(i,4).text())
                except:
                    unit_cost = 0

                u = pgso_procurement_data.update().where(pgso_procurement_data.c.id == id).\
                    values(
                    description = description,
                    quantity = quantity,
                    unit = unit,
                    unit_cost = unit_cost
                    )
                conn.execute(u)

            u = pgso_procurement.update().where(pgso_procurement.c.id == po_id).values(status='po')
            conn.execute(u)

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Inserted")
            msg.setInformativeText('Data Inserted to the Database')
            msg.setWindowTitle("PGSO Purchase Order")
            msg.exec_()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Something Went Wrong , Do it Again')
            msg.setWindowTitle("Error")
            msg.exec_()
        self.tabWidget.setCurrentIndex(2)

    def gen_cancel_button_action(self):
        global gen_tab
        self.tabWidget.setCurrentIndex(gen_tab)

    ##po
    def show_po_name(self):
        self.po_name.clear()
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_department = sqc.Database().pgso_department
        s = pgso_department.select().where(pgso_department.c.type == self.po_type.currentText())
        s_value = conn.execute(s)
        for val in s_value:
            self.po_name.addItem(val[2])

    po_dict = {}
    def show_po_list(self):
        self.po_list.clear()
        self.po_dict = {}
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_department = sqc.Database().pgso_department
        s = pgso_department.select().where(pgso_department.c.type == self.po_type.currentText()). \
            where(pgso_department.c.name == self.po_name.currentText())
        s_value = conn.execute(s)
        id = 0
        for val in s_value:
            id = val[0]

        pgso_procurement = sqc.Database().pgso_procurement
        s = pgso_procurement.select().where(pgso_procurement.c.department_id == id).where(pgso_procurement.c.status == 'po')
        s_value = conn.execute(s)

        for val in s_value:
            self.po_dict.update({'{}_(#{})_{}'.format(self.po_name.currentText(),val[0],val[2]) : val[0]})

        for key in self.po_dict.keys():
            self.po_list.addItem(key)

    def po_generate_po_action(self):
        excel = win32com.client.DispatchEx('Excel.Application')
        scbk = excel.Workbooks.Open(os.getcwd() +'\\template\\PO.xlsx',ReadOnly = True)
        scsht = scbk.Worksheets[1]
        excel.Visible = True
        lrow = find_cell(scsht, '', search_order='by_row', search_sheet='basic')['row']
        lcol = find_cell(scsht, '', search_order='by_col', search_sheet='basic')['col']
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_procurement_data = sqc.Database().pgso_procurement_data
        s = pgso_procurement_data.select()
        s_value = conn.execute(s)
        for val in s_value:
            range_insert = scsht.Range("A15:L15")
            range_insert.EntireRow.Insert()

        s_value = conn.execute(s)
        i = 1
        row = 15
        for val in s_value:
            scsht.Cells(row,1).Value = i
            scsht.Cells(row,2).Value = val[3]
            scsht.Cells(row,3).Value = val[2]
            scsht.Cells(row,4).Value = val[1]
            scsht.Cells(row, 4).HorizontalAlignment = xlwings.constants.HAlign.xlHAlignLeft
            scsht.Cells(row, 4).WrapText = False
            scsht.Cells(row,9).Value = val[4]
            scsht.Cells(row,11).Formula = '=C{} * I{}'.format(row,row)
            row += 1
            i += 1
        p = inflect.engine()
        date_index = find_cell(scsht,'Date:',lrow=lrow,lcol=lcol)
        scsht.Cells(date_index['row'],date_index['col'] + 1).Value = datetime.datetime.now()
        scsht.Cells(date_index['row'],date_index['col'] + 1).HorizontalAlignment = xlwings.constants.HAlign.xlHAlignCenter
        scsht.Cells(date_index['row'], date_index['col'] + 1).Columns.AutoFit()
        purpose_remarks = find_cell(scsht,'Conforme:',lrow=3000,lcol=lcol)
        scsht.Range('I'+ str(purpose_remarks['row'] -4)).Formula = '=SUM(K16:K{})'.format(purpose_remarks['row']-1)
        scsht.Range('A'+ str(purpose_remarks['row'] -4)).Value = p.number_to_words(float(str(scsht.Range('I'+ str(purpose_remarks['row'] -4)).Value)))

    def po_generate_pr_action(self):
        excel = win32com.client.DispatchEx('Excel.Application')
        scbk = excel.Workbooks.Open(os.getcwd() +'\\template\\PR.xlsx',ReadOnly = True)
        scsht = scbk.Worksheets[1]
        excel.Visible = True
        lrow = find_cell(scsht, '', search_order='by_row', search_sheet='basic')['row']
        lcol = find_cell(scsht, '', search_order='by_col', search_sheet='basic')['col']
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_procurement_data = sqc.Database().pgso_procurement_data
        s = pgso_procurement_data.select()
        s_value = conn.execute(s)
        for val in s_value:
            range_insert = scsht.Range("A16:J16")
            range_insert.EntireRow.Insert()

        s_value = conn.execute(s)
        i = 1
        row = 16
        for val in s_value:
            scsht.Cells(row,1).Value = i
            scsht.Cells(row,2).Value = val[3]
            scsht.Cells(row,3).Value = val[1]
            scsht.Cells(row, 3).HorizontalAlignment = xlwings.constants.HAlign.xlHAlignLeft
            scsht.Cells(row,8).Value = val[2]
            scsht.Cells(row,9).Value = val[4]
            scsht.Cells(row,10).Formula = '=H{} * I{}'.format(row,row)
            row += 1
            i += 1

        department_index = find_cell(scsht,'Department : ',lrow=lrow,lcol=lcol)
        purpose_remarks = find_cell(scsht,'Purpose/Remarks :',lrow=3000,lcol=lcol)
        scsht.Cells(department_index['row'],department_index['col'] + 2).Value = self.po_name.currentText()
        scsht.Cells(purpose_remarks['row'],purpose_remarks['col']).Value = 'Purpose/Remarks : For {} Use.'.format(self.po_type.currentText())
        scsht.Cells(purpose_remarks['row'],purpose_remarks['col'] + 2).Value = 'For {} use.'.format(self.po_name.currentText())
        scsht.Cells(purpose_remarks['row'] -1,purpose_remarks['col'] + 8).Formula = '=SUM(J16:J{})'.format(purpose_remarks['row']-2)

    def po_edit_action(self):
        global gen_tab
        gen_tab = 5
        self.tabWidget.setCurrentIndex(4)
        temp = self.po_list.currentItem().text()
        temp2 = temp.split('(')[0] +temp.split(')')[1]
        self.gen_title.setText(temp2)
        po_id = self.po_dict[self.po_list.currentItem().text()]
        engine = sqc.Database().engine
        conn = engine.connect()
        pgso_procurement_data = sqc.Database().pgso_procurement_data
        s = pgso_procurement_data.select().where(pgso_procurement_data.c.po_id == po_id).order_by(pgso_procurement_data.c.description)
        s_value = conn.execute(s)
        table = self.gen_table_widget
        table.setRowCount(0)
        for val in s_value :
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[3])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 3, QTableWidgetItem(str(val[2])))
            table.setItem(row_position, 4, QTableWidgetItem(str(val[4])))
            table.resizeColumnsToContents()
            table.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)

    def po_delete_action(self):
        try:
            po_id = self.po_dict[self.po_list.currentItem().text()]
            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_procurement = sqc.Database().pgso_procurement
            s = pgso_procurement.delete().where(pgso_procurement.c.id == po_id)
            conn.execute(s)
            pgso_procurement_data = sqc.Database().pgso_procurement_data
            s = pgso_procurement_data.delete().where(pgso_procurement_data.c.po_id == po_id)
            conn.execute(s)
            os.remove(os.getcwd() + '\\excel\\' + str(self.po_dict[self.po_list.currentItem().text()]) + '.xlsx')
            self.show_po_list()
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Success")
            msg.setInformativeText('File Deleted Properly')
            msg.setWindowTitle("Information")
            msg.exec_()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('Something went wrong in file deletion')
            msg.setWindowTitle("Error")
            msg.exec_()



    ##settings
    def show_settings(self):
        self.settings_account_table.setRowCount(0)
        engine = sqc.Database().engine
        pgso_admin = sqc.Database().pgso_admin
        conn= engine.connect()
        #admin_table
        s = pgso_admin.select().order_by(asc(pgso_admin.c.username))
        s_value = conn.execute(s)
        table = self.settings_account_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))

        self.settings_department_table.setRowCount(0)
        pgso_department = sqc.Database().pgso_department
        # admin_table
        s = pgso_department.select().order_by(asc(pgso_department.c.type))
        s_value = conn.execute(s)
        table = settings_department_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))

    def settings_edit_account_action(self, table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            username = table.item(r, 1).text()
            password = table.item(r, 2).text()
            ad = Accounts_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id, username, password, operationType='edit')
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def settings_add_account_action(self):
        try:
            ad = Accounts_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id, '', '', operationType='add')
        except:
            pass

    def settings_delete_account_action(self, table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_admin = sqc.Database().pgso_admin
            s = pgso_admin.delete().where(pgso_admin.c.userid == id)
            conn.execute(s)
            conn.close()
            self.show_settings()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def settings_edit_department_action(self, table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            type = table.item(r, 1).text()
            name = table.item(r, 2).text()
            ad = Department_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id, type, name, operationType='edit')
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def settings_add_department_action(self):
        try:
            ad = Department_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id, '', '', operationType='add')
        except:
            pass

    def settings_delete_department_action(self, table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_department = sqc.Database().pgso_department
            s = pgso_department.delete().where(pgso_department.c.id == id)
            conn.execute(s)
            conn.close()
            self.show_settings()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    #price_list
    def show_price_list(self):
        global price_list_table
        price_list_table.setRowCount(0)
        engine = sqc.Database().engine
        pgso_price_list = sqc.Database().pgso_price_list
        conn= engine.connect()
        #admin_table
        s = pgso_price_list.select().order_by(asc(pgso_price_list.c.item))
        s_value = conn.execute(s)
        table = price_list_table
        for val in s_value:
            row_position = table.rowCount()
            table.insertRow(row_position)
            table.setItem(row_position, 0, QTableWidgetItem(str(val[0])))
            table.setItem(row_position, 1, QTableWidgetItem(str(val[1])))
            table.setItem(row_position, 2, QTableWidgetItem(str(val[2])))
            table.resizeColumnsToContents()
        conn.close()

    def menu_price_list_action(self):
        self.tabWidget.setCurrentIndex(7)
        self.show_price_list()

    def price_list_edit_action(self, table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            item = table.item(r, 1).text()
            price = table.item(r, 2).text()
            ad = Item_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id, item, price, operationType='edit')
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

    def price_list_add_action(self):
        try:
            ad = Item_Dialogue(self)
            ad.show()
            ad.ShowDialogue(id, '', '', operationType='add')
        except:
            pass

    def price_list_delete_action(self, table):
        try:
            r = table.currentRow()
            id = table.item(r, 0).text()
            engine = sqc.Database().engine
            conn = engine.connect()
            pgso_price_list = sqc.Database().pgso_price_list
            s = pgso_price_list.delete().where(pgso_price_list.c.id == id)
            conn.execute(s)
            conn.close()
            self.show_price_list()
        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText('No Rows Selected')
            msg.setWindowTitle("Error")
            msg.exec_()

def main():
    app=QApplication(sys.argv)
    window=MainApp()
    window.show()
    app.exec_()

if __name__=='__main__':
    main()
