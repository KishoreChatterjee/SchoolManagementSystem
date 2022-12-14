from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
import sys
from PyQt5.uic import loadUiType
import mysql.connector as con
from datetime import date
from PyQt5.QtPrintSupport import QPrinter,QPrintDialog
import pandas as pd

ui, _ = loadUiType('Scms.ui')

class MainApp(QMainWindow, ui):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)

        self.tabWidget.setCurrentIndex(0)
        self.tabWidget.tabBar().setVisible(False)
        self.menubar.setVisible(False)

        self.b01.clicked.connect(self.login)
        self.menu11.triggered.connect(self.add_new_student_tab)
        self.b11.clicked.connect(self.save_student_details)
        self.menu12.triggered.connect(self.edit_or_delete_student_tab)
        self.cb21.currentIndexChanged.connect(self.fill_details_when_comboBox_selected)
        self.b21.clicked.connect(self.edit_student_details)
        self.b22.clicked.connect(self.delete_student_details)

        self.menu21.triggered.connect(self.marks_student_tab)
        self.b31.clicked.connect(self.save_marks_details)
        self.cb32.currentIndexChanged.connect(self.fill_exam_name_in_ComboBox_for_selected_registration_number)
        self.b32.clicked.connect(self.fill_exam_details_in_textBox_for_selected_exam_name)
        self.b33.clicked.connect(self.update_marks_details)
        self.b34.clicked.connect(self.delete_marks_details)

        self.menu31.triggered.connect(self.Attendence_student_tab)
        self.b41.clicked.connect(self.save_Attendence_details)
        self.cb42.currentIndexChanged.connect(self.fill_date_in_ComboBox_for_regno_selected)
        self.b42.clicked.connect(self.fill_attendance_status_on_button_clicked)
        self.b43.clicked.connect(self.update_attendance_details)
        self.b44.clicked.connect(self.delete_attendance_details)

        self.menu41.triggered.connect(self.fees_student_tab)
        self.b51.clicked.connect(self.save_fees_details)

        self.b81.clicked.connect(self.print_file)
        self.b82.clicked.connect(self.cancel_print)
        self.cb52.currentIndexChanged.connect(self.fill_reciept_details_in_TextBox_for_reciept_combo_selected)
        self.b52.clicked.connect(self.update_fees_details)
        self.b53.clicked.connect(self.delete_fees_details)

        self.menu51.triggered.connect(self.show_report)
        self.menu52.triggered.connect(self.show_report)
        self.menu53.triggered.connect(self.show_report)
        self.menu54.triggered.connect(self.show_report)

        self.menu61.triggered.connect(self.logout)
        self.exp.clicked.connect(self.export_table)


########login form########
    def login(self):
        un = self.tb01.text()
        pw = self.tb02.text()

        if (un == "admin" and pw == "admin"):
            self.menubar.setVisible(True)
            self.tabWidget.setCurrentIndex(1)
        else:
            QMessageBox.information(self,"School Management System","Invalid Admin login Details,Try again!")
            self.l01.setText("Invalid Admin login Details,Try again!")

    def add_new_student_tab(self):
        self.tabWidget.setCurrentIndex(2)
        self.fill_next_registration_number()

    def edit_or_delete_student_tab(self):
        self.tabWidget.setCurrentIndex(3)
        self.fill_registration_number_in_ComboBox()

    def marks_student_tab(self):
        self.tabWidget.setCurrentIndex(4)
        self.fill_registration_number_in_ComboBox_in_mark_tab()

    def Attendence_student_tab(self):
        self.tabWidget.setCurrentIndex(5)
        self.fill_registration_number_in_ComboBox_in_attendance_tab()
        self.tb41.setText(str(date.today()))

    def fees_student_tab(self):
        self.tabWidget.setCurrentIndex(6)
        self.fill_registration_number_in_ComboBox_in_attendance_tab()
        self.tb41.setText(str(date.today()))
        self.fill_registration_number_in_ComboBox_in_fees_tab()
        self.fill_next_reciept_number()
        self.tb54.setText(str(date.today()))
        self.fill_reciept_number_in_ComboBox_in_edit_fees_tab()

###### Filling next registration number#########

    def fill_next_registration_number(self):
        try:
            rn = 0
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from student")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    rn += 1
            self.tb11.setText(str(rn+1))
        except con.Error as e:
            print("Error occured in select student reg number"+ e)

####### save Student details function########

    def save_student_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            registration_number = self.tb11.text()
            full_name = self.tb12.text()
            gender = self.cb11.currentText()
            date_of_birth = self.tb13.text()
            age = self.tb14.text()
            address = self.mtb11.text()
            phone = self.tb15.text()
            email = self.tb16.text()
            standard = self.cb12.currentText()

            qry = "insert into student (registration_number,full_name,gender,date_of_birth,age,address,phone,email,standard) values(%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            value = (registration_number,full_name,gender,date_of_birth,age,address,phone,email,standard)
            cursor.execute(qry,value)
            mydb.commit()

            self.l11.setText("Students details saved successfully")
            QMessageBox.information(self,"School Management System","Students details saved successfully")
            self.tb11.setText("")
            self.tb12.setText("")
            self.tb13.setText("")
            self.tb14.setText("")
            self.mtb11.setText("")
            self.tb15.setText("")
            self.tb16.setText("")
            self.tabWidget.setCurrentIndex(1)
        except con.Error as e:
            self.l11.setText("Error in save student detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in save student detail form")

#########Filling registration number in combo box###########

    def fill_registration_number_in_ComboBox(self):
        try:
            self.cb21.clear()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from student")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    self.cb21.addItem(str(stud[1]))

        except con.Error as e:
            print("Error occured in filling student reg number in combo box"+ e)

##########Fills details when combobox is selected#########

    def fill_details_when_comboBox_selected(self):
        try:
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from student where registration_number = '"+ self.cb21.currentText() +"'")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    self.tb21.setText(str(stud[2]))
                    self.tb21_2.setText(str(stud[3]))
                    self.tb22.setText(str(stud[4]))
                    self.tb23.setText(str(stud[5]))
                    self.mtb21.setText(str(stud[6]))
                    self.tb24.setText(str(stud[7]))
                    self.tb25.setText(str(stud[8]))
                    self.tb26.setText(str(stud[9]))

        except con.Error as e:
            print("Error occured in filling details when combobox is selected."+ e)

#########Edit student details function##########

    def edit_student_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            registration_number = self.cb21.currentText()
            full_name = self.tb21.text()
            gender = self.tb21_2.text()
            date_of_birth = self.tb22.text()
            age = self.tb23.text()
            address = self.mtb21.text()
            phone = self.tb24.text()
            email = self.tb25.text()
            standard = self.tb26.text()

            qry = "update student set full_name = '"+ full_name +"', gender = '"+ gender +"', date_of_birth = '"+ date_of_birth +"',age = '"+ age +"', address = '"+ address +"', phone = '"+ phone +"', email = '"+ email +"', standard = '"+ standard +"' where registration_number = '"+ registration_number +"'"
            cursor.execute(qry)
            mydb.commit()

            self.l21.setText("Students details updated successfully")
            QMessageBox.information(self,"School Management System","Students details updated successfully")
            self.tb21.setText("")
            self.tb21_2.setText("")
            self.tb22.setText("")
            self.tb23.setText("")
            self.mtb21.setText("")
            self.tb24.setText("")
            self.tb25.setText("")
            self.tb26.setText("")
            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:
            self.l21.setText("Error in editing student detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in updating student detail form")

##########Delete student detail function#########

    def delete_student_details(self):
        m = QMessageBox.question(self,"Delete","Are you sure you want to delete this student details?",QMessageBox.Yes|QMessageBox.No)
        if (m == QMessageBox.Yes):
            try:
                mydb = con.connect(host="localhost", user="root", password="", db="school")
                cursor = mydb.cursor()
                registration_number = self.cb21.currentText()
                qry = "delete from student where registration_number = '" + registration_number + "'"
                cursor.execute(qry)
                mydb.commit()
                self.l21.setText("Students details deleted successfully")
                QMessageBox.information(self, "School Management System", "Students details deleted successfully")
                self.tabWidget.setCurrentIndex(1)
            except con.Error as e:
                self.l21.setText("Error in delete student detail form" + e)
                QMessageBox.information(self, "School Management System", "Error in deleting student detail form")

########## marks Coding ##############

    def fill_registration_number_in_ComboBox_in_mark_tab(self):
        try:
            self.cb31.clear()
            self.cb32.clear()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from student")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    self.cb31.addItem(str(stud[1]))
                    self.cb32.addItem(str(stud[1]))


        except con.Error as e:
            print("Error occured in filling student reg number in combo box"+ e)

    #####save marks coding#########

    def save_marks_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            registration_number = self.cb31.currentText()
            exam_name = self.tb31.text()
            english = self.tb33.text()
            language = self.tb32.text()
            math = self.tb34.text()
            science = self.tb35.text()
            social = self.tb36.text()


            qry = "insert into marks (registration_number,exam_name,english,language,math,science,social) values(%s,%s,%s,%s,%s,%s,%s)"
            value = (registration_number,exam_name,english,language,math,science,social)
            cursor.execute(qry,value)
            mydb.commit()

            self.l31.setText("marks details added successfully")
            QMessageBox.information(self,"School Management System","marks details added successfully")
            self.tb31.setText("")
            self.tb32.setText("")
            self.tb33.setText("")
            self.tb34.setText("")
            self.tb35.setText("")
            self.tb36.setText("")
            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:
            self.l31.setText("Error in adding student marks detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in adding student marks detail form")

    def fill_exam_name_in_ComboBox_for_selected_registration_number(self):
        try:
            self.cb33.clear()
            registration_number = self.cb32.currentText()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from marks where registration_number = '"+ registration_number +"'")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    self.cb33.addItem(str(stud[2]))

        except con.Error as e:
            print("Error occured in filling exam name in combo box"+ e)

    def fill_exam_details_in_textBox_for_selected_exam_name(self):
        try:
            registration_number = self.cb32.currentText()
            exam_name = self.cb33.currentText()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from marks where registration_number = '"+ registration_number +"' and exam_name = '"+ exam_name +"'")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    self.tb37.setText(str(stud[4]))
                    self.tb38.setText(str(stud[3]))
                    self.tb39.setText(str(stud[5]))
                    self.tb310.setText(str(stud[6]))
                    self.tb311.setText(str(stud[7]))

        except con.Error as e:
            print("Error occured in filling marks details in text box"+ e)

    #####edit marks coding#########

    def update_marks_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            registration_number = self.cb32.currentText()
            exam_name = self.cb33.currentText()

            language = self.tb37.text()
            english = self.tb38.text()
            math = self.tb39.text()
            science = self.tb310.text()
            social = self.tb311.text()

            qry = "update marks set english = '"+ english +"', language = '"+ language +"', math = '"+ math +"',science = '"+ science +"', social = '"+ social +"' where registration_number = '"+ registration_number +"' and exam_name = '"+ exam_name +"'"
            cursor.execute(qry)
            mydb.commit()

            self.l32.setText("Marks details updated successfully")
            QMessageBox.information(self,"School Management System","marks details updated successfully")
            self.tb37.setText("")
            self.tb38.setText("")
            self.tb39.setText("")
            self.tb310.setText("")
            self.tb311.setText("")
            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:
            self.l32.setText("Error in editing marks detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in updating marks detail form")

    #####delete marks coding#########

    def delete_marks_details(self):
        m = QMessageBox.question(self,"Delete","Are you sure you want to delete this student details?",QMessageBox.Yes|QMessageBox.No)
        if (m == QMessageBox.Yes):
            try:
                mydb = con.connect(host="localhost", user="root", password="", db="school")
                cursor = mydb.cursor()
                registration_number = self.cb32.currentText()
                exam_name = self.cb33.currentText()

                qry = "delete from marks where registration_number = '" + registration_number + "' and exam_name = '"+ exam_name +"'"
                cursor.execute(qry)
                mydb.commit()
                self.l32.setText("Marks details deleted successfully")
                QMessageBox.information(self, "School Management System", "marks details deleted successfully")
                self.tb37.setText("")
                self.tb38.setText("")
                self.tb39.setText("")
                self.tb310.setText("")
                self.tb311.setText("")
                self.tabWidget.setCurrentIndex(1)
            except con.Error as e:
                self.l32.setText("Error in delete marks detail form" + e)
                QMessageBox.information(self, "School Management System", "Error in deleting marks detail form")

#####Attendence coding#########

    def fill_registration_number_in_ComboBox_in_attendance_tab(self):
        try:
            self.cb41.clear()
            self.cb42.clear()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from student")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    self.cb41.addItem(str(stud[1]))
                    self.cb42.addItem(str(stud[1]))

        except con.Error as e:
            print("Error occured in filling student reg number in combo box"+ e)

    #####save Attendence coding#########

    def save_Attendence_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            registration_number = self.cb41.currentText()
            attendance_date = self.tb41.text()
            morning = self.tb42.text()
            evening = self.tb43.text()

            qry = "insert into attendance (registration_number,attendance_date,morning,evening) values(%s,%s,%s,%s)"
            value = (registration_number,attendance_date,morning,evening)
            cursor.execute(qry,value)
            mydb.commit()

            self.l41.setText("Attendance details added successfully")
            QMessageBox.information(self,"School Management System","Attendance details added successfully")

            self.tb42.setText("")
            self.tb43.setText("")

        except con.Error as e:
            self.l41.setText("Error in adding student Attendance detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in adding student attendance detail form")

    def fill_date_in_ComboBox_for_regno_selected(self):
        try:
            self.cb43.clear()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from attendance where registration_number = '"+ self.cb42.currentText() +"'")
            result = cursor.fetchall()
            if result:
                for att in result:
                    self.cb43.addItem(str(att[2]))


        except con.Error as e:
            print("Error occured in filling student reg number in combo box"+ e)

    def fill_attendance_status_on_button_clicked(self):
        try:
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from attendance where registration_number = '"+ self.cb42.currentText() +"' and attendance_date = '"+ self.cb43.currentText() +"'")
            result = cursor.fetchall()
            if result:
                for att in result:
                    self.tb44.setText(str(att[3]))
                    self.tb45.setText(str(att[4]))

        except con.Error as e:
            print("Error occured in filling attendance status"+ e)

    #####edit Attendence coding#########

    def update_attendance_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            registration_number = self.cb42.currentText()
            attendance_date = self.cb43.currentText()
            morning = self.tb44.text()
            evening = self.tb45.text()

            qry = "update attendance set morning = '"+ morning +"', evening = '"+ evening +"' where registration_number = '"+ registration_number +"' and attendance_date = '"+ attendance_date +"'"
            cursor.execute(qry)
            mydb.commit()
            self.tb44.setText("")
            self.tb45.setText("")

            self.l42.setText("Attendance details updated successfully")
            QMessageBox.information(self,"School Management System","Attendance details updated successfully")
        except con.Error as e:
            self.l42.setText("Error in editing attendance detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in updating attendance detail form")

    #####delete Attendence coding#########

    def delete_attendance_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            registration_number = self.cb42.currentText()
            attendance_date = self.cb43.currentText()


            qry = "delete from attendance where registration_number = '"+ registration_number +"' and attendance_date = '"+ attendance_date +"'"
            cursor.execute(qry)
            mydb.commit()
            self.tb44.setText("")
            self.tb45.setText("")

            self.l42.setText("Attendance details deleted successfully")
            QMessageBox.information(self,"School Management System","Attendance details deleted successfully")
            self.tabWidget.setCurrentIndex(1)
        except con.Error as e:
            self.l42.setText("Error in deleting attendance detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in deleting attendance detail form")

######### fees tab coding##########

    def fill_registration_number_in_ComboBox_in_fees_tab(self):
        try:
            self.cb51.clear()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from student")
            result = cursor.fetchall()
            if result:
                for stud in result:
                    self.cb51.addItem(str(stud[1]))

        except con.Error as e:
            print("Error occured in filling student reg number in combo box"+ e)

    def fill_next_reciept_number(self):
        try:
            rn = 0
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from fees")
            result = cursor.fetchall()
            if result:
                for rec in result:
                    rn += 1
            self.tb51.setText(str(rn+1))

        except con.Error as e:
            print("Error occured in filling student reciept number in combo box"+ e)

    def save_fees_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            reciept_number = self.tb51.text()
            registration_number = self.cb51.currentText()
            reason = self.tb52.text()
            amount = self.tb53.text()
            fees_date = self.tb54.text()

            qry = "insert into fees (reciept_number,registration_number,reason,amount,fees_date) values(%s,%s,%s,%s,%s)"
            value = (reciept_number,registration_number,reason,amount,fees_date)
            cursor.execute(qry,value)
            mydb.commit()

            self.l51.setText("Fees details added successfully")
            QMessageBox.information(self,"School Management System","Fees details added successfully")

            self.fill_reciept_number_in_ComboBox_in_edit_fees_tab()

            self.l81.setText(self.tb51.text())
            self.l82.setText(self.tb54.text())
            self.l86.setText(self.tb54.text())
            self.l84.setText(self.tb53.text())
            self.l85.setText(self.tb52.text())
            cursor.execute("select * from student where registration_number = '"+ registration_number +"'")
            result = cursor.fetchone()
            if result:
                self.l83.setText(str(result[2]))
            self.tabWidget.setCurrentIndex(8)
            self.tb52.setText("")
            self.tb53.setText("")
            self.fill_next_reciept_number()

        except con.Error as e:
            self.l51.setText("Error in adding student fees detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in adding student fees detail form")

    def fill_reciept_number_in_ComboBox_in_edit_fees_tab(self):
        try:
            self.cb52.clear()
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from fees")
            result = cursor.fetchall()
            if result:
                for rec in result:
                    self.cb52.addItem(str(rec[1]))

        except con.Error as e:
            print("Error occured in filling student reciept number in combo box"+ e)

    def fill_reciept_details_in_TextBox_for_reciept_combo_selected(self):
        try:
            mydb = con.connect(host="localhost",user="root",password="",db="school")
            cursor = mydb.cursor()
            cursor.execute("select * from fees where reciept_number = '"+ self.cb52.currentText() +"'")
            result = cursor.fetchall()
            if result:
                for rec in result:
                    self.tb55.setText(str(rec[2]))
                    self.tb56.setText(str(rec[3]))
                    self.tb57.setText(str(rec[4]))
                    self.tb58.setText(str(rec[5]))

        except con.Error as e:
            print("Error occured in filling student reciept details in Text box"+ e)

    def update_fees_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            reciept_number = self.cb52.currentText()
            registration_number = self.tb55.text()
            reason = self.tb56.text()
            amount = self.tb57.text()
            fees_date = self.tb58.text()
            qry = "update fees set registration_number = '"+ registration_number +"', reason = '"+ reason +"' , amount = '"+ amount +"', fees_date = '"+ fees_date +"'where reciept_number = '"+ reciept_number +"'"
            cursor.execute(qry)
            mydb.commit()
            self.l52.setText("Fees details updated successfully")
            QMessageBox.information(self,"School Management System","Fees details updated successfully")

            self.l81.setText(self.cb52.currentText())
            self.l82.setText(self.tb58.text())
            self.l86.setText(self.tb58.text())
            self.l84.setText(self.tb57.text())
            self.l85.setText(self.tb56.text())
            cursor.execute("select * from student where registration_number = '"+ registration_number +"'")
            result = cursor.fetchone()
            if result:
                self.l83.setText(str(result[2]))
            self.tabWidget.setCurrentIndex(8)

            self.tb55.setText("")
            self.tb56.setText("")
            self.tb57.setText("")
            self.tb58.setText("")

        except con.Error as e:
            self.l52.setText("Error in editing fees detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in updating fees detail form")

    def delete_fees_details(self):
        try:
            mydb = con.connect(host="localhost", user="root", password="", db="school")
            cursor = mydb.cursor()
            reciept_number = self.cb52.currentText()
            qry = "delete from fees where reciept_number = '"+ reciept_number +"'"
            cursor.execute(qry)
            mydb.commit()
            self.l52.setText("Fees details deleted successfully")
            QMessageBox.information(self,"School Management System","Fees details deleted successfully")

            self.tb55.setText("")
            self.tb56.setText("")
            self.tb57.setText("")
            self.tb58.setText("")
            self.tabWidget.setCurrentIndex(1)

        except con.Error as e:
            self.l52.setText("Error in deleting fees detail form"+ e)
            QMessageBox.information(self, "School Management System", "Error in deleting fees detail form")

    ###########print Report Form###########

    def show_report(self):
        table_name = self.sender()
        self.l61.setText(table_name.text())
        self.tabWidget.setCurrentIndex(7)
        try:
            self.tableReport.setRowCount(0)
            print(table_name.text())
            if (table_name.text() == "Students Reports"):
                mydb = con.connect(host="localhost", user="root", password="", db="school")
                cursor = mydb.cursor()
                qry = "select registration_number,full_name,gender,date_of_birth,age,address,phone,email,standard from student"
                cursor.execute(qry)
                result = cursor.fetchall()
                r = 0
                c = 0
                for row_number,row_data in enumerate(result):
                    r += 1
                    c = 0
                    for row_number, data in enumerate(row_data):
                        c += 1
                self.tableReport.clear()
                self.tableReport.setColumnCount(c)
                for row_number, row_data in enumerate(result):
                    self.tableReport.insertRow(row_number)
                    for column_number, data in enumerate(row_data):
                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                self.tableReport.setHorizontalHeaderLabels(['Register_Number','Name','Gender','Date_of_Birth','Age','Address','phone','Email','Standard'])


            if (table_name.text() == "Marks Reports"):
                mydb = con.connect(host="localhost", user="root", password="", db="school")
                cursor = mydb.cursor()
                qry = "select registration_number,exam_name,english,language,math,science,social from marks"
                cursor.execute(qry)
                result = cursor.fetchall()
                r = 0
                c = 0
                for row_number,row_data in enumerate(result):
                    r += 1
                    c = 0
                    for row_number, data in enumerate(row_data):
                        c += 1
                self.tableReport.clear()
                self.tableReport.setColumnCount(c)
                for row_number, row_data in enumerate(result):
                    self.tableReport.insertRow(row_number)
                    for column_number, data in enumerate(row_data):
                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                self.tableReport.setHorizontalHeaderLabels(['Register_Number','Exam_Name','English','Language','Math','Science','Social'])

            if (table_name.text() == "Attendence Reports"):
                mydb = con.connect(host="localhost", user="root", password="", db="school")
                cursor = mydb.cursor()
                qry = "select registration_number,attendance_date,morning,evening from attendance"
                cursor.execute(qry)
                result = cursor.fetchall()
                r = 0
                c = 0
                for row_number,row_data in enumerate(result):
                    r += 1
                    c = 0
                    for row_number, data in enumerate(row_data):
                        c += 1
                self.tableReport.clear()
                self.tableReport.setColumnCount(c)
                for row_number, row_data in enumerate(result):
                    self.tableReport.insertRow(row_number)
                    for column_number, data in enumerate(row_data):
                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                self.tableReport.setHorizontalHeaderLabels(['Register_Number','Attendance_date','Morning','Evening'])

            if (table_name.text() == "Fees Reports"):
                mydb = con.connect(host="localhost", user="root", password="", db="school")
                cursor = mydb.cursor()
                qry = "select reciept_number,registration_number,reason,amount,fees_date from fees"
                cursor.execute(qry)
                result = cursor.fetchall()
                r = 0
                c = 0
                for row_number,row_data in enumerate(result):
                    r += 1
                    c = 0
                    for row_number, data in enumerate(row_data):
                        c += 1
                self.tableReport.clear()
                self.tableReport.setColumnCount(c)
                for row_number, row_data in enumerate(result):
                    self.tableReport.insertRow(row_number)
                    for column_number, data in enumerate(row_data):
                        self.tableReport.setItem(row_number, column_number, QTableWidgetItem(str(data)))
                self.tableReport.setHorizontalHeaderLabels(['Reciept_number','Registration_number','Reason','Amount','Fees_date'])

        except con.Error as e:
            print(" Error in report form " + e)

 ########## Export report to excel#######

    def export_table(self):
        print("Export started")
        columnHeaders = []
        for j in range(self.tableReport.model().columnCount()):
            columnHeaders.append(self.tableReport.horizontalHeaderItem(j).text())
            print(columnHeaders)

        df = pd.DataFrame(columns = columnHeaders)
        for row in range(self.tableReport.rowCount()):
            for col in range(self.tableReport.columnCount()):
                df.at[row,columnHeaders[col]] = self.tableReport.item(row,col).text()

        xlname = self.l61.text() + str(date.today()) + " .xlsx"
        df.to_excel(xlname,index=False)
        QMessageBox.information(self, "School Management System", "File exported successfully! " + xlname)

    ###########print file function###########

    def print_file(self):
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer,self)
        if dialog.exec_() == QPrintDialog.Accepted:
            self.tabWidget.print_(printer)

    def cancel_print(self):
        self.tabWidget.setCurrentIndex(1)

#######logout function#######

    def logout(self):
        try:
            self.menubar.setVisible(False)
            self.tb01.setText("")
            self.tb02.setText("")
            self.tabWidget.setCurrentIndex(0)
            QMessageBox.information(self, "School Management System", "logout successfully")
        except:
            QMessageBox.information(self, "School Management System", "error in logging out")


def main():
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()

if __name__ == '__main__':
    main()


