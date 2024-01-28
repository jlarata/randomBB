
RandomBB is a program to choose randomly a book from your library for you to read.
If the book has already been read, RandomBB will choose another for you.
Also RandomBB will inform you the "already been read" list of books that were choosen, in case any of them is not really read and you prefer to read it now.

RandomBB uses a database named "mydata" created and mantained with Shelves. 
This database contains a list of objects "Libro" extracted with openpyxl 3.1.2 from an excel spreadsheet (biblioteca.xlsx)
Commented lines below 120 contains the method to create and store that database. They're not needed once the database is created and stored, but can be useful if using a new spreadsheet, or if expanded, or using csv instead, etc.

"dist" and "build" folders are used by pyinstaller 5.10.1 to create a simple .exe
For the .exe in the "dist" folder to work, you need to manually copy the 3 mydata.db files.
