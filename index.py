# PyQt5 used to link to .ui modules
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

# Used to read and write to .xlsx files
from xlsxwriter import *
from xlrd import *

import datetime             # Used to obtain the current date/time
import MySQLdb              # Connects the python and .ui to the Database
import sys

# Loads the main UI of the application
from PyQt5.uic import loadUiType

# Calls the UIs
ui, _ = loadUiType('library.ui')
login, _ = loadUiType('login.ui')

# Handles the Login Processes and UI
class Login (QWidget, login):
    
    # Loads the login UI
    def __init__(self):
        QWidget.__init__(self)
        self.setupUi(self)

        # Logs the user in on button click
        self.pushButton.clicked.connect(self.Handle_Login)

    # Handles the Login process
    def Handle_Login(self):

        # Connect to the MySQL Database
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        
        # Calls a cursor on the DB to access the data
        self.cur = self.db.cursor()

        # Get inserted information from the GUI as processable text
        username = self.lineEdit.text()
        password = self.lineEdit_2.text()

        # SQL query to be executed in MySQL
        sql = ''' SELECT * FROM users'''

        # Executes the SQL command
        self.cur.execute(sql)

        # Stores the result of the executes SQL query
        data = self.cur.fetchall()
        
        # If no users in DB, allow login by default
        if data == ():

            # Displays confirmation if login successful
            info = QMessageBox.information(self, 'Login Successful','Login Successful!', QMessageBox.Ok)
            
            # If 'Ok' is clicked, performs the specified actions
            if info == QMessageBox.Ok:

                # Open the Main App Window if login successful
                self.window2 = MainApp()
                self.close()                    # Close the login window
                self.window2.show()             # Shows the MainApp window
        
        # Iterates through the dat
        for row in data:

            # Enables login via username and email both
            if row[1] == username or row[2] == username:

                # Checks if the password is correct
                if row[3] ==  password:

                    # Gives a success message
                    info = QMessageBox.information(self, 'Login Successful','Login Successful!', QMessageBox.Ok)
                    if info == QMessageBox.Ok:

                        # Opens the Main App window
                        self.window2 = MainApp()
                        self.close()
                        self.window2.show()
                
                else:
                    warning = QMessageBox.warning(self, 'Incorrect Details','Please enter the correct login details.', QMessageBox.Ok)
                    if warning == QMessageBox.Ok:
                        self.lineEdit_2.setText('')

# Handles the Main Application UI
class MainApp (QMainWindow, ui):
    
    # Handles all the main functions after application loading
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)

        # Calling the functions required on application startup
        self.Handle_UI_Changes()
        self.Handle_buttons()
        
        # Shows the associated data in the respective tables
        self.Show_Operations()
        self.Show_Category()
        self.Show_Publisher()
        self.Show_Author()
        self.Show_Books()
        self.Show_Clients()

        # Shows the associated data in the respective comboboxes
        self.Combobox_Author()
        self.Combobox_Category()
        self.Combobox_Publisher()

    # Handles the UI changes
    def Handle_UI_Changes(self):
        self.Hide_Themes()
        self.tabWidget.tabBar().setVisible(False)

    # Handles the buttons
    def Handle_buttons(self):
        
        # Handle the Operations
        self.pushButton_3.clicked.connect(self.Handle_Operations)
        
        # Show/Hide Themes
        self.pushButton_8.clicked.connect(self.Show_Themes)
        self.pushButton_21.clicked.connect(self.Hide_Themes)

        # Toggle between the various themes
        self.pushButton_18.clicked.connect(self.aqua)
        self.pushButton_19.clicked.connect(self.breezedark)
        self.pushButton_20.clicked.connect(self.breezelight)
        self.pushButton_27.clicked.connect(self.classic)
        self.pushButton_28.clicked.connect(self.darkblue)
        self.pushButton_31.clicked.connect(self.ubuntu)

        # Navigate between tabs
        self.pushButton.clicked.connect(self.Open_Operations)
        self.pushButton_2.clicked.connect(self.Open_Books)
        self.pushButton_26.clicked.connect(self.Open_Clients)
        self.pushButton_6.clicked.connect(self.Open_Users)
        self.pushButton_7.clicked.connect(self.Open_Settings)

        # Add New Author, Publisher, Category
        self.pushButton_14.clicked.connect(self.Add_Author)
        self.pushButton_15.clicked.connect(self.Add_Publisher)
        self.pushButton_16.clicked.connect(self.Add_Category)

        # Delete an existing Author, Publisher, Category
        self.pushButton_23.clicked.connect(self.Delete_Author)
        self.pushButton_24.clicked.connect(self.Delete_Publisher)
        self.pushButton_25.clicked.connect(self.Delete_Category)

        # Book related operations
        self.pushButton_4.clicked.connect(self.Add_New_Book)
        self.pushButton_9.clicked.connect(self.Search_Book)
        self.pushButton_5.clicked.connect(self.Edit_Book)
        self.pushButton_10.clicked.connect(self.Delete_Book)

        # Client related operations
        self.pushButton_17.clicked.connect(self.Add_Client)
        self.pushButton_33.clicked.connect(self.Search_Client)
        self.pushButton_34.clicked.connect(self.Edit_Client)
        self.pushButton_35.clicked.connect(self.Delete_Client)

        # User related operations
        self.pushButton_11.clicked.connect(self.Add_Users)
        self.pushButton_12.clicked.connect(self.Login)
        self.pushButton_13.clicked.connect(self.Edit_User)
        self.pushButton_22.clicked.connect(self.Delete_User)

        # Export operations
        self.pushButton_36.clicked.connect(self.Export_Operations)
        self.pushButton_29.clicked.connect(self.Export_Books)
        self.pushButton_30.clicked.connect(self.Export_Clients)

    # Shows the themes tab
    def Show_Themes(self):
        self.groupBox_6.show()

    # Hides the themes tab
    def Hide_Themes(self):
        self.groupBox_6.hide()

    ################# Toggle between various tabs via buttons ################# 
    # Uses the Tab Indices to switch between tabs
    def Open_Operations(self):
        self.tabWidget.setCurrentIndex(0)

    def Open_Books(self):
        self.tabWidget.setCurrentIndex(1)

    def Open_Clients(self):
        self.tabWidget.setCurrentIndex(2)

    def Open_Users(self):
        self.tabWidget.setCurrentIndex(3)

    def Open_Settings(self):
        self.tabWidget.setCurrentIndex(4)

    ################# Book Operations #################
    # Adds a new book
    def Add_New_Book(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_2.text()
        book_desc = self.textEdit.toPlainText()
        book_category = self.comboBox_3.currentIndex()
        book_author = self.comboBox_4.currentIndex()
        book_publisher = self.comboBox_5.currentIndex()
        book_price = self.lineEdit_4.text()

        self.cur.execute(''' 
            INSERT INTO book (book_name, book_desc, category, author, publisher, price) 
            VALUES (%s, %s, %s, %s, %s, %s)
        ''', (book_title, book_desc, book_category, book_author, book_publisher, book_price))

        self.db.commit()

        # Shows a confirmation message in the status bar
        self.statusBar().showMessage("New Book ({0}) Inserted.".format(book_title))

        # Resets the respcetive fields once entry is done
        self.lineEdit_2.setText('')
        self.textEdit.setPlainText('')
        self.comboBox_3.setCurrentIndex(0)
        self.comboBox_4.setCurrentIndex(0)
        self.comboBox_5.setCurrentIndex(0)
        self.lineEdit_4.setText('')
        
        # Updates the 'View Books' tab to show the recently added book(s)
        self.Show_Books()

    # Search for an existing book in the DB
    def Search_Book(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        book_title = self.lineEdit_8.text()

        # If search bar empty, display message
        if (book_title == ''):
            self.statusBar().showMessage("No Book Found.")

        # Execute the search
        else:
            sql = ''' SELECT * FROM book where book_name = %s'''
            self.cur.execute(sql, [(book_title)])

            # Fetch only one entry from the database
            data = self.cur.fetchone()

            # Returns the formatted and processed data in the respective fields
            self.lineEdit_6.setText(data[1])
            self.textEdit_2.setPlainText(data[2])
            self.lineEdit_7.setText(str(data[0]))
            self.comboBox_7.setCurrentIndex(data[3])
            self.comboBox_6.setCurrentIndex(data[4])
            self.comboBox_8.setCurrentIndex(data[5])
            self.lineEdit_5.setText(str(data[6]))

            self.statusBar().showMessage("Search Result Retuned.")

    # Edit details for an existing book
    def Edit_Book(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        search_book_title = self.lineEdit_8.text()

        book_id = self.lineEdit_7.text()
        book_title = self.lineEdit_8.text()
        book_desc = self.textEdit_2.toPlainText()
        book_category = self.comboBox_7.currentIndex()
        book_author = self.comboBox_6.currentIndex()
        book_publisher = self.comboBox_6.currentIndex()
        book_price = self.lineEdit_5.text()
    
        self.cur.execute('''
            UPDATE book SET book_id=%s, book_name=%s, book_desc=%s, category=%s, author=%s, publisher=%s, price=%s WHERE book_name = %s
        ''', (book_id, book_title, book_desc, str(book_category), str(book_author), str(book_publisher), str(book_price), search_book_title))
        
        self.db.commit()
        self.statusBar().showMessage("Book data ({0}) updated.".format(search_book_title))

        self.Show_Books()

    # Delete an existing book
    def Delete_Book(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        search_book_title = self.lineEdit_8.text()

        warning = QMessageBox.question(self, 'Delete Book','Are you sure you want to delete this book? ({0})'.format(search_book_title), QMessageBox.Yes | QMessageBox.No)
        
        # Asks for confirmation before deleting
        if (warning == QMessageBox.Yes):
            sql = ''' DELETE FROM book WHERE book_name = %s '''
            self.cur.execute(sql, [(search_book_title )])
            self.db.commit()
            QMessageBox.information(self, 'Book Deleted','Book deleted successfully!', QMessageBox.Ok)

        self.Show_Books()

    ################# Client Operations #################
    # Add a new Client to the DB
    def Add_Client(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        first_name = self.lineEdit_3.text()
        middle_name = self.lineEdit_23.text()
        last_name = self.lineEdit_25.text()
        client_email = self.lineEdit_20.text()
        phone = self.lineEdit_26.text()

        self.cur.execute('''
            INSERT INTO clients (first_name, middle_name, last_name, client_email, phone) VALUES (%s, %s, %s, %s, %s)
        ''', (first_name, middle_name, last_name, client_email, phone))

        self.db.commit()
        self.db.close()
        QMessageBox.information(self, 'New Client Created','New Client created successfully!', QMessageBox.Ok)

        self.Show_Clients()

    # Search for an existing client
    def Search_Client(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        first_name = self.lineEdit_34.text()
        last_name = self.lineEdit_35.text()

        self.cur.execute('''
            SELECT * FROM clients WHERE first_name = %s AND last_name = %s
        ''',(first_name, last_name))

        data = self.cur.fetchone()
        
        # Enables the group box after a successful search
        self.groupBox_10.setEnabled(True)

        self.lineEdit_32.setText(data[1])
        self.lineEdit_36.setText(data[2])
        self.lineEdit_37.setText(data[3])
        self.lineEdit_33.setText(data[4])
        self.lineEdit_38.setText(str(data[5]))

        self.statusBar().showMessage("Search Result Retuned.")

    # Edit details for an existing client
    def Edit_Client(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        search_fname = self.lineEdit_34.text()
        search_lname = self.lineEdit_35.text()
        first_name = self.lineEdit_32.text()
        middle_name = self.lineEdit_36.text()
        last_name = self.lineEdit_37.text()        
        client_email = self.lineEdit_33.text()
        phone = self.lineEdit_38.text()

        self.cur.execute(''' 
            UPDATE clients SET first_name = %s, middle_name = %s, last_name = %s, client_email = %s, phone = %s WHERE first_name = %s AND last_name = %s
        ''', (first_name, middle_name, last_name, client_email, phone, search_fname, search_lname))

        self.db.commit()
        QMessageBox.information(self, 'Edit(s) Successful','Details updated successfully!', QMessageBox.Ok)

        self.lineEdit_32.setText('')
        self.lineEdit_36.setText('')
        self.lineEdit_37.setText('')
        self.lineEdit_33.setText('')
        self.lineEdit_38.setText('')

        # Disables the group box after a successful edit
        self.groupBox_10.setEnabled(False)
        self.Show_Clients()

    # Delete an exiting client
    def Delete_Client(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        first_name = self.lineEdit_34.text()
        last_name = self.lineEdit_35.text()

        warning = QMessageBox.question(self, 'Delete Client','Are you sure you want to delete this client? ({0})'.format(
            first_name+' '+last_name), QMessageBox.Yes | QMessageBox.No)

        if (warning == QMessageBox.Yes):
            self.cur.execute(''' DELETE FROM clients  WHERE first_name = %s AND last_name = %s''', (first_name, last_name))

            self.db.commit()
            QMessageBox.information(self, 'Client Deleted','Client deleted successfully!', QMessageBox.Ok)

            self.lineEdit_32.setText('')
            self.lineEdit_36.setText('')
            self.lineEdit_37.setText('')
            self.lineEdit_33.setText('')
            self.lineEdit_38.setText('')

            self.groupBox_10.setEnabled(False)
        
        self.Show_Clients()

    ################# User Operations #################
    # Add a new user
    def Add_Users(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        username = self.lineEdit_9.text()
        email = self.lineEdit_10.text()
        password = self.lineEdit_12.text()
        password2 = self.lineEdit_11.text()

        # Adds the user only if both passwords are same
        if password == password2:
            self.cur.execute(''' INSERT INTO users (username, user_email, user_pwd) VALUES (%s, %s, %s)
            ''', (username, email, password))

            self.db.commit()
            QMessageBox.information(self, 'User Created','New user successfully created!', QMessageBox.Ok)
        
        # Throws error if both passwords not same
        else:
            warning = QMessageBox.warning(self, 'Password','Please enter same password in both fields.', QMessageBox.Ok)

            # Clears the password fields
            if warning == QMessageBox.Ok:
                self.lineEdit_12.setText('')
                self.lineEdit_11.setText('')
    
    # Login feature for the user to be able to edit their information
    def Login(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        username = self.lineEdit_13.text()
        password = self.lineEdit_14.text()

        sql = ''' SELECT * FROM users'''

        self.cur.execute(sql)

        # Fetches all the rows from the DB
        data = self.cur.fetchall()
        
        for row in data:
            if row[1] == username:
                if row[3] ==  password:

                    # Enable the group box after successful login
                    self.groupBox_7.setEnabled(True)
                    QMessageBox.information(self, 'Login Successful','Login Successful!', QMessageBox.Ok)

                    self.lineEdit_17.setText(row[1])
                    self.lineEdit_15.setText(row[2])                
                
                else:
                    warning = QMessageBox.warning(self, 'Incorrect Details','Please enter the correct login details.', QMessageBox.Ok)
                    if warning == QMessageBox.Ok:
                        self.lineEdit_14.setText('')
        
        self.lineEdit_14.setText('')

    # Edit existing user information    
    def Edit_User(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        username = self.lineEdit_13.text()
        email = self.lineEdit_15.text()
        password = self.lineEdit_18.text()
        password2 = self.lineEdit_16.text()

        if password == password2:
            self.cur.execute(''' 
                UPDATE users SET user_email = %s, user_pwd = %s WHERE username = %s
            ''', (email, password, username))

            self.db.commit()
            QMessageBox.information(self, 'Edit(s) Successful','Details updated successfully!', QMessageBox.Ok)

            self.lineEdit_17.setText('')
            self.lineEdit_15.setText('')
            self.lineEdit_18.setText('')
            self.lineEdit_16.setText('')
            self.lineEdit_13.setText('')

            self.groupBox_7.setEnabled(False)

        else:
            warning = QMessageBox.warning(self, 'Password','Please enter same password in both fields.', QMessageBox.Ok)

            if warning == QMessageBox.Ok:
                self.lineEdit_18.setText('')
                self.lineEdit_16.setText('')

    # Delete existing user
    def Delete_User(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        username = self.lineEdit_17.text()
        password = self.lineEdit_18.text()
        password2 = self.lineEdit_16.text()

        if password == password2:
            warning = QMessageBox.question(self, 'Delete User','Are you sure you want to delete this user? ({0})'.format(username), QMessageBox.Yes | QMessageBox.No)
            if (warning == QMessageBox.Yes):
                self.cur.execute(''' DELETE FROM users  WHERE username = %s ''', (username,))

                self.db.commit()
                QMessageBox.information(self, 'User Deleted','User deleted successfully!', QMessageBox.Ok)

                self.lineEdit_17.setText('')
                self.lineEdit_15.setText('')
                self.lineEdit_18.setText('')
                self.lineEdit_16.setText('')

                self.groupBox_7.setEnabled(False)

        else:
            warning = QMessageBox.warning(self, 'Password','Please enter same password in both fields.', QMessageBox.Ok)

            if warning == QMessageBox.Ok:
                self.lineEdit_18.setText('')
                self.lineEdit_16.setText('')


    ################# Settings Operations #################
    # Add a new Author
    def Add_Author(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        auth_name = self.lineEdit_19.text()

        self.cur.execute('''
            INSERT INTO author (auth_name) VALUES (%s)
        ''', (auth_name,))

        self.db.commit()
        self.statusBar().showMessage("New Author ({0}) Inserted.".format(auth_name))

        self.lineEdit_19.setText('')
        self.Show_Author()
        self.Combobox_Author()
    
    # Add a new Publisher
    def Add_Publisher(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        pub_name = self.lineEdit_22.text()

        self.cur.execute('''
            INSERT INTO publisher (pub_name) VALUES (%s)
        ''', (pub_name,))

        self.db.commit()
        self.statusBar().showMessage("New Publisher ({0}) Inserted.".format(pub_name))
        
        self.lineEdit_22.setText('')
        self.Show_Publisher()
        self.Combobox_Publisher()

    # Add a new category
    def Add_Category(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        cat_name = self.lineEdit_24.text()

        self.cur.execute('''
            INSERT INTO category (cat_name) VALUES (%s)
        ''', (cat_name,))

        self.db.commit()
        self.statusBar().showMessage("New Category ({0}) Inserted.".format(cat_name))

        self.lineEdit_24.setText('')
        self.Show_Category()
        self.Combobox_Category()

    # Delete an existing author
    def Delete_Author(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        auth_name = self.lineEdit_19.text()

        warning = QMessageBox.question(self, 'Delete Author','Are you sure you want to delete this author? ({0})'.format(auth_name), QMessageBox.Yes | QMessageBox.No)
        if (warning == QMessageBox.Yes):
            self.cur.execute(''' DELETE FROM author WHERE auth_name = %s ''', (auth_name,))

            self.db.commit()
            QMessageBox.question(self, 'Author Deleted','Author deleted successfully!', QMessageBox.Ok)

            self.lineEdit_19.setText('')
            self.Show_Author()
            self.Combobox_Author()

    # Delete an existing publisher
    def Delete_Publisher(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        pub_name = self.lineEdit_22.text()

        warning = QMessageBox.question(self, 'Delete Publisher','Are you sure you want to delete this publisher? ({0})'.format(pub_name), QMessageBox.Yes | QMessageBox.No)
        if (warning == QMessageBox.Yes):
            self.cur.execute(''' DELETE FROM publisher WHERE pub_name = %s ''', (pub_name,))

            self.db.commit()
            QMessageBox.question(self, 'Publisher Deleted','Publisher deleted successfully!', QMessageBox.Ok)

            self.lineEdit_22.setText('')
            self.Show_Publisher()
            self.Combobox_Publisher()

    # Delete an existing category
    def Delete_Category(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        cat_name = self.lineEdit_24.text()

        warning = QMessageBox.question(self, 'Delete Category','Are you sure you want to delete this category? ({0})'.format(cat_name), QMessageBox.Yes | QMessageBox.No)
        if (warning == QMessageBox.Yes):
            self.cur.execute(''' DELETE FROM category WHERE cat_name = %s ''', (cat_name,))

            self.db.commit()
            QMessageBox.question(self, 'Category Deleted','Category deleted successfully!', QMessageBox.Ok)

            self.lineEdit_24.setText('')
            self.Show_Category()
            self.Combobox_Category()

    ################# Operation Functions #################
    # Handles the Day-to-Day functioning of the library
    def Handle_Operations(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        book_title = self.lineEdit.text()
        first_name = self.lineEdit_27.text()
        last_name = self.lineEdit_21.text()
        action = self.comboBox.currentText()
        day = self.comboBox_2.currentIndex()
        date = datetime.datetime.now()
        date = date.strftime("%Y-%m-%d %H:%M:%S")

        self.cur.execute(''' INSERT INTO operations(book_id, client_id, operation, days, date)
            VALUES ((SELECT book_id FROM book WHERE book_name = %s), (
                SELECT client_id FROM clients WHERE first_name = %s AND last_name = %s), %s, %s, %s);
        ''', (book_title, first_name, last_name, action, day, date))

        self.db.commit()
        self.statusBar().showMessage('Operation Added')

        self.Show_Operations()

    ################# Show Data in Tables #################
    # Shows all Authors in the database
    def Show_Author(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT * FROM author''')
        data = self.cur.fetchall()

        if data:
            self.tableWidget_2.setRowCount(0)
            self.tableWidget_2.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_2.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(row_position)

    # Shows all publishers in the database
    def Show_Publisher(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT * FROM publisher''')
        data = self.cur.fetchall()

        if data:
            self.tableWidget_3.setRowCount(0)
            self.tableWidget_3.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_3.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_3.rowCount()
                self.tableWidget_3.insertRow(row_position)
    
    # Shows all categories in the database
    def Show_Category(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT * FROM category''')
        data = self.cur.fetchall()

        if data:
            self.tableWidget_4.setRowCount(0)
            self.tableWidget_4.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_4.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_4.rowCount()
                self.tableWidget_4.insertRow(row_position)

    # Shows all the info about the Books in the database
    def Show_Books(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT book_id, book_name, auth_name, cat_name, pub_name, price FROM book INNER JOIN category, publisher, author 
                            WHERE book.category = category.cat_id AND book.author = author.auth_id AND book.publisher = publisher.pub_id;''')
        data = self.cur.fetchall()

        if data:
            self.tableWidget_5.setRowCount(0)
            self.tableWidget_5.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_5.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_5.rowCount()
                self.tableWidget_5.insertRow(row_position)

    # Shows all the clients in the database
    def Show_Clients(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT * FROM clients''')
        data = self.cur.fetchall()

        if data:
            self.tableWidget_6.setRowCount(0)
            self.tableWidget_6.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget_6.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget_6.rowCount()
                self.tableWidget_6.insertRow(row_position)

    # Shows all the operations of the library
    def Show_Operations(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT op_id, book.book_name, clients.first_name, clients.middle_name, clients.last_name, operation, days, date 
            FROM operations INNER JOIN book, clients 
            WHERE operations.book_id = book.book_id AND operations.client_id = clients.client_id
        ''')

        data = self.cur.fetchall()

        if data:
            self.tableWidget.setRowCount(0)
            self.tableWidget.insertRow(0)
            for row, form in enumerate(data):
                for column, item in enumerate(form):
                    self.tableWidget.setItem(row, column, QTableWidgetItem(str(item)))
                    column += 1
                row_position = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row_position)

    # Shows the authors in the database in the combobox for easier use
    def Combobox_Author(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT auth_name FROM author ''')
        data = self.cur.fetchall()
        
        # Clears the combo box to prevent repetition of values after addition of new values
        self.comboBox_4.clear()
        self.comboBox_6.clear()

        # Adds '<none>' to the combobox for entering the data more consistently  in the DB
        self.comboBox_4.addItem('<none>')
        self.comboBox_6.addItem('<none>')
        for author in data:
            for i in author:
                self.comboBox_4.addItem(i)
                self.comboBox_6.addItem(i)

    # Shows the publishers in the database in the combobox for easier use
    def Combobox_Publisher(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT pub_name FROM publisher ''')
        data = self.cur.fetchall()
        
        self.comboBox_5.clear()
        self.comboBox_8.clear()
        self.comboBox_5.addItem('<none>')
        self.comboBox_8.addItem('<none>')
        for publisher in data:
            for i in publisher:
                self.comboBox_5.addItem(i)
                self.comboBox_8.addItem(i)

    # Shows the categories in the database in the combobox for easier use
    def Combobox_Category(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute(''' SELECT cat_name FROM category ''')
        data = self.cur.fetchall()
        
        self.comboBox_3.clear()
        self.comboBox_7.clear()
        self.comboBox_3.addItem('<none>')
        self.comboBox_7.addItem('<none>')
        for category in data:
            for i in category:
                self.comboBox_3.addItem(i)
                self.comboBox_7.addItem(i)

    ################# Export functions ################
    # Exports the operations data into the 'operations.xlsx' file
    def Export_Operations(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT op_id, book.book_name, clients.first_name, clients.middle_name, clients.last_name, operation, days, date 
            FROM operations INNER JOIN book, clients 
            WHERE operations.book_id = book.book_id AND operations.client_id = clients.client_id
        ''')

        data = self.cur.fetchall()
        
        # Creates a workbook
        wb = Workbook('operations.xlsx')
        
        # Adds a worksheet to the workbook
        sheet1 = wb.add_worksheet()

        # Create the header to better understand the exported file
        sheet1.write(0,0,'Operation ID')
        sheet1.write(0,1,'Book Title')
        sheet1.write(0,2,'Client FName')
        sheet1.write(0,3,'Client MName')
        sheet1.write(0,4,'Client LName')
        sheet1.write(0,5,'Operation')
        sheet1.write(0,6,'Days')
        sheet1.write(0,7,'Date')

        #### Iterate through 'data' and add a new row one by one
        # Writes the file from row-1 since row-0 has the respective headers
        row_number = 1
        for row in data:
            column_number = 0

            for item in row:
                sheet1.write(row_number, column_number, item)
                column_number += 1
            row_number += 1

        # Closes the open file
        wb.close()
        
        # Export completion confirmation
        QMessageBox.information(self, 'Export Complete','Data exported to operations.xlsx')

    # Exports the books data into the 'books.xlsx' file
    def Export_Books(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT book_id, book_name, auth_name, book_desc, cat_name, pub_name, price FROM book INNER JOIN category, publisher, author 
            WHERE book.category = category.cat_id AND book.author = author.auth_id AND book.publisher = publisher.pub_id;''')

        data = self.cur.fetchall()

        wb = Workbook('books.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0,'Book ID')
        sheet1.write(0,1,'Book Title')
        sheet1.write(0,2,'Author')
        sheet1.write(0,3,'Description')
        sheet1.write(0,4,'Category')
        sheet1.write(0,5,'Publisher')
        sheet1.write(0,6,'Price')

        row_number = 1
        for row in data:
            column_number = 0

            for item in row:
                sheet1.write(row_number, column_number, item)
                column_number += 1
            row_number += 1

        wb.close()
        
        QMessageBox.information(self, 'Export Complete','Data exported to books.xlsx')

    # Exports the clients data into the 'clients.xlsx' file
    def Export_Clients(self):
        self.db = MySQLdb.connect(host='localhost', user='root', password='2187', db='librarysys')
        self.cur = self.db.cursor()

        self.cur.execute('''SELECT * FROM clients''')
        data = self.cur.fetchall()

        wb = Workbook('clients.xlsx')
        sheet1 = wb.add_worksheet()

        sheet1.write(0,0,'Client ID')
        sheet1.write(0,1,'FName')
        sheet1.write(0,2,'MName')
        sheet1.write(0,3,'LName')
        sheet1.write(0,4,'Email')
        sheet1.write(0,5,'Phone')

        row_number = 1
        for row in data:
            column_number = 0

            for item in row:
                sheet1.write(row_number, column_number, item)
                column_number += 1
            row_number += 1

        wb.close()
        
        QMessageBox.information(self, 'Export Complete','Data exported to clients.xlsx')

    ################# Apply Themes ################
    # Aqua Theme
    def aqua(self):
        style = open('themes/aqua.css', 'r')        # Open the CSS syle sheet as read only
        style = style.read()                        # Reads the style sheet
        self.setStyleSheet(style)                   # Set the style sheet for the application UI

    # Breeze Dark Theme
    def breezedark(self):
        style = open('themes/breezedark.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    # Breeze Light Theme    
    def breezelight(self):
        style = open('themes/breezelight.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    # Classic Theme    
    def classic(self):
        style = open('themes/classic.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    # Dark Blue Theme    
    def darkblue(self):
        style = open('themes/darkblue.css', 'r')
        style = style.read()
        self.setStyleSheet(style)

    # Ubuntu Theme
    def ubuntu(self):
        style = open('themes/ubuntu.css', 'r')
        style = style.read()
        self.setStyleSheet(style)
    
# Driver code for the application
def main():
    app = QApplication(sys.argv)

    # Loads the Login page by default
    window = Login()
    window.show()
    app.exec()

# Calling the main function
if __name__ == "__main__":
    main()