# InventoryManagementSoftware-DepoTakipYazilimi
Inventory Management Software for ISTAÇ A.Ş.


![MainScreen](https://user-images.githubusercontent.com/80919382/140296498-88e2331e-767b-4d6f-9382-d26e8326a857.PNG)

# Features of Inventory Management Program

The program consists of three different parts: Inventory List, Registered Product List and Junk List.

Inventory List:
- Products are in inventory, not belong to someone yet. 
- Properties of product (columns of inventory table) are product id, product name, amount of product, measurement unit and explanation of product.
- Functions of inventory list are adding nonregistered product, deleting nonregistered product, registering product to someone, editing nonregistered product, filtering products and exporting the list as excel file.

Registered Product List:
- Products belong to some employees in this list.
- Properties of product (columns of registered product table) are product id, product name, amount of product, measurement unit, employee who gives product to someone, employee who receives the product, date of the operation and explanation of product.
- Functions of registered product list are deleting registered product, taking product from employee back to inventory, editing registered product, filtering products (by their name or by searching employee name who gives or recieves the product) and exporting the list as excel file.

Junk List:
- This list keeps records of deleted product.
- Properties of deleted product (columns of junk table) are product id, product name, amount of product, measurement unit and explanation of the reason why the product is deleted.
- Functions of junk list are filtering products, exporting the list as excel file and cleaning the table (removing all items of table).

