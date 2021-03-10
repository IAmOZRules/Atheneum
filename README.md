# Library-Managament-System
A library management system made using Python, MySQL and PyQt5


## Needed python packages:

##### To Install, run ```pip install <package name>``` in CMD or Terminal.
```
PyQt5
mysqlclient
xlxswriter
xlrd
```
> Run ```index.py``` to open the application.

### How to run the application on the desktop:

1. Download/Clone this repository.

To Clone the repository, if you have git installed, run the following command in a Terminal/CMD:
```git clone https://github.com/IAmOZRules/Library-Management-System.git```

2. Create a new database named ```librarysys``` in MySQL.

------- OR -------

2. Open the ```index.py``` file, replace ```'librarysys'``` everywhere in the code with whatever your DB is named.

**For example**: If you named your DB to ```'library'``` replace  ```'librarysys'``` with ```'library'``` everywhere in the ```index.py``` file.

3. In the ```index.py``` file, change ```password = '2187'``` to whatever your MySQL Server password is!

**For example**: If your MySQL Server password  ```'helloworld'``` replace  ```password = '2187'``` with ```password = 'helloworld'``` everywhere in the ```index.py``` file.

4. Run the ```database.sql``` file in MySQL workbench in the new created schema.
5. Make sure all the required python packages are installed!
6. Navigate to the folder in which all the files are present.
7. Run the ```index.py``` file.
8. For first time login, use username = ```system``` and password = ```system```.
9. Make sure to create a user for you after opening.
10. Figure out the extra stuff yourself!

Do **NOT** skip any step, or the app won't run.

So much for an *easy* install smh.
