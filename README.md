# InventoryManagementSoftware-DepoTakipYazilimi
Inventory Management Software for ISTAÇ A.Ş.

Download as windows application: https://drive.google.com/file/d/1WabOUUamUYsVMFBMRMcd64geoBok6N8H/view?usp=sharing
Please ReadMe.txt after extracting 'Windows Application' file in the .zip file.


![MainScreen](https://user-images.githubusercontent.com/80919382/140296498-88e2331e-767b-4d6f-9382-d26e8326a857.PNG)

# Features of Inventory Management Program

The program consists of three different parts: Inventory List, Registered Product List and Junk List.

Inventory List:
- Products are in inventory, not belong to someone yet. 
- Properties of product (columns of inventory table) are product id, product name, amount of product, measurement unit and explanation of product.
- Functions of inventory list are adding nonregistered product, deleting nonregistered product, registering product to someone, editing nonregistered product, filtering products and exporting the list as excel file.


![MainScreen](https://user-images.githubusercontent.com/80919382/140299537-688877b7-08ef-4044-bf41-be7fc35c42fb.PNG)


Registered Product List:
- Products belong to some employees in this list.
- Properties of product (columns of registered product table) are product id, product name, amount of product, measurement unit, employee who gives product to someone, employee who receives the product, date of the operation and explanation of product.
- Functions of registered product list are deleting registered product, taking product from employee back to inventory, editing registered product, filtering products (by their name or by searching employee name who gives or recieves the product) and exporting the list as excel file.


![RegisteredProduct](https://user-images.githubusercontent.com/80919382/140299557-69e36ee6-6f13-40ae-8ddf-4ac21a942222.PNG)



Junk List:
- This list keeps records of deleted product.
- Properties of deleted product (columns of junk table) are product id, product name, amount of product, measurement unit and explanation of the reason why the product is deleted.
- Functions of junk list are filtering products, exporting the list as excel file and cleaning the table (removing all items of table).


![Junk](https://user-images.githubusercontent.com/80919382/140299581-47a57049-14db-4899-b96f-46115a8c25aa.PNG)



# General Information about the Program

- Files must be placed as they are on the GitHub when you want to download and try the program.
- Tkinter and pandas libraries are mainly used.
- Program keeps log records. Logs.log file must be in the same directory as .py file. Contents of log records are date of operation, person who made the operation (this info comes from username of computer), list name, operation name, properties of product which is in the process.
- Program uses excel file as save files. When you want to use these excel files which are located in 'Tablolar' file, please use copies of them. Do not change original excel files.
- Program runs online at ISTAC AS. normally. I used their local network to save the data files. Every computer which is connected to their server can use the program at the same time. They can see the changes which is done by another computer. This version is local version of the program.
