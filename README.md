# RandomBB
_(a program i designed to choose randomly a book to read from my library)_.

 I already had an excel spreadsheet to keep track of my books so the idea was to use that spreadsheet as raw data.

RandomBB can import the data (using openpyxl 3.1.2) and create a persistent database using Python Shelves. In this case, the db contains a simple list of objects named "Libro" (book) with the expected attributes (title, author, etc) plus a boolean "isLeido" (has been read).

Every time is executed randomBB checks if the xlsx file has new records and updates the database.

Its easy to manipulate and gives some options to the user.

You can ignore "dist" and "build" folders. They are used by pyinstaller 5.10.1 library to create a windows executable.
For the .exe in the "dist" folder to work the first time its required to manually copy the 3 mydata.db files. 
