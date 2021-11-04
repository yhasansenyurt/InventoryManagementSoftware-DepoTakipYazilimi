# This program is made by Hasan Senyurt for ISTAC A.S. - Inventory Management Software

from tkinter import *
import pandas as pd
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime
import getpass

############################################################
# Arranging row colors of the tables.
def coloring3():
    for i in table3.get_children():
        if int(table3.index(i)) % 2 == 0:
            table3.item(i,tags=("even"))
        else:
            table3.item(i,tags=("odd"))          
    
    table3.tag_configure("even",background="Azure2") 
    table3.tag_configure("odd",background="ghost white")

def coloring2():
    for i in table2.get_children():
        if int(table2.index(i)) % 2 == 0:
            table2.item(i,tags=("even"))
        else:
            table2.item(i,tags=("odd"))          
    
    table2.tag_configure("even",background="Azure2") 
    table2.tag_configure("odd",background="ghost white")
    
def coloring():
    for i in table.get_children():
        if int(table.index(i)) % 2 == 0:
            table.item(i,tags=("even"))
        else:
            table.item(i,tags=("odd"))          
    
    table.tag_configure("even",background="Azure2") 
    table.tag_configure("odd",background="ghost white")

############################################################


############################################################
# Keeping log records.
def logs():
    global logfile
    
    logfile = open("logs.log","a")
    logfile.write(str(datetime.now())[:19]+"\t"+format(getpass.getuser(),"20s"))

def logsclose():
    logfile.close()
############################################################

############################################################
# Saving changed data before quiting program.
def beforeExit():
    if messagebox.askokcancel("Çıkış", "Çıkış yapmak istiyor musunuz?"):
        
        if filter_cancel['state'] == ACTIVE:
            returnTable()
        
        if filter_cancel2['state'] == ACTIVE:
            returnTable2()

        if filter_cancel3['state'] == ACTIVE:
            returnTable3()
        
        export("local")
        export2("local")
        export3("local")
        window.destroy()
############################################################

############################################################
# Saving changes on table to data file after every operation.
def updateTable(tablo,liste):
    
        global p,p2,p3

        if p != 1 and liste == "Depo Listesi.xlsx":
            for i in tablo.get_children():
                tablo.delete(i)

            data = pd.read_excel("Tablolar/"+liste)

            data = data.fillna('')

            for i in range(0,len(data)):
                tablo.insert(parent = '', index='end',id=i+1, values=[data.loc[i][j] for j in range(0,len(data.columns))])
            coloring()
                

        if p2 != 1 and liste == "Zimmet Listesi.xlsx":
            for i in tablo.get_children():
                tablo.delete(i)

            data = pd.read_excel("Tablolar/"+liste)

            data = data.fillna('')

            for i in range(0,len(data)):
                tablo.insert(parent = '', index='end',id=i+1, values=[data.loc[i][j] for j in range(0,len(data.columns))])
            coloring2()

        if p3 != 1 and liste == "Hurda Listesi.xlsx":
            for i in tablo.get_children():
                tablo.delete(i)

            data = pd.read_excel("Tablolar/"+liste)

            data = data.fillna('')

            for i in range(0,len(data)):
                tablo.insert(parent = '', index='end',id=i+1, values=[data.loc[i][j] for j in range(0,len(data.columns))])
            coloring3()

############################################################        


############################################################
# Loading data to both three tables from data files.
def importExcel3():
    
    global items3
    items3 = pd.read_excel("Tablolar/Hurda Listesi.xlsx")
    

    global columns3

    columns3 = list()

    for i in items3.columns:
        columns3.append(i)
   
    table3['columns'] = columns3

    table3.column("#0", width=0, stretch=NO)

    for i in range(0,len(columns3)):
        if(i ==4):
            table3.column(columns3[i],width=428,minwidth=750) #392
        elif i==0:
            table3.column(columns3[i],width=100,minwidth=100)
        elif i == 1:
            table3.column(columns3[i],width=300,minwidth=300)

        else:
            table3.column(columns3[i],width=100,minwidth=100)

    table3.heading("#0", text="")

    for i in range(0,len(columns3)):
        table3.heading(columns3[i],text=columns3[i],anchor="w")

    items3 = items3.fillna('')
    items3['Açıklama'] = items3['Açıklama'].astype(str)
    
    for i in range(0,len(items3)):
        table3.insert(parent = '', index='end',id=i+1, values=[items3.loc[i][j] for j in range(0,len(items3.columns))])
    coloring3()
        
def importExcel2():

    global items2
    items2 = pd.read_excel("Tablolar/Zimmet Listesi.xlsx")
    

    global columns2

    columns2 = list()

    for i in items2.columns:
        columns2.append(i)
   
    table2['columns'] = columns2

    table2.column("#0", width=0, stretch=NO)

    for i in range(0,len(columns2)):
        if(i==1):
            table2.column(columns2[i],width=260,minwidth=260)
        elif i ==7:
            table2.column(columns2[i],width=126,minwidth=750)
        elif i==0:
            table2.column(columns2[i],width=90,minwidth=90)
        elif i ==5 or i ==6:
            table2.column(columns2[i],width=134,minwidth=134)
        elif i ==4:
            table2.column(columns2[i],width=134,minwidth=134)
        else:
            table2.column(columns2[i],width=75,minwidth=75)

    table2.heading("#0", text="")

    for i in range(0,len(columns2)):
        table2.heading(columns2[i],text=columns2[i],anchor="w")

    items2 = items2.fillna('')

    items2['Tarih'] = items2['Tarih'].astype(str)
    items2['Açıklama'] = items2['Açıklama'].astype(str)


    for i in range(0,len(items2)):
        table2.insert(parent = '', index='end',id=i+1, values=[items2.loc[i][j] for j in range(0,len(items2.columns))])
    coloring2()

def importExcel():
    
    global items

    items = pd.read_excel("Tablolar/Depo Listesi.xlsx")

    global columns

    columns = list()

    for i in items.columns:
        columns.append(i)
   
    table['columns'] = columns

    table.column("#0", width=0, stretch=NO)

    for i in range(0,len(columns)):
        if(i ==4):
            table.column(columns[i],width=378,minwidth=750) #392
        elif i==0:
            table.column(columns[i],width=100,minwidth=100)
        elif i == 1:
            table.column(columns[i],width=300,minwidth=300)

        else:
            table.column(columns[i],width=75,minwidth=75)

    table.heading("#0", text="")

    for i in range(0,len(columns)):
        table.heading(columns[i],text=columns[i],anchor="w")

    items = items.fillna('')
    items['Açıklama'] = items['Açıklama'].astype(str)

    for i in range(0,len(items)):
        table.insert(parent = '', index='end',id=i+1, values=[items.loc[i][j] for j in range(0,len(items.columns))])

    coloring()
############################################################

############################################################
# Solving problem because of upper 'i' character in Turkish alphabet.
def entryUpper(entry):
    
    up = str()
    for i in entry:
        if i == 'i':
            up = up + 'İ'
        else:
            up = up + i.upper()
    return up

############################################################

############################################################
# Clearing all items on junk table.
def removeAll():
    for i in table3.get_children():
        table3.delete(i)

    logs()
    logfile.write("HURDA LİSTESİ"+"\t"+"TEMİZLEME"+"      "+"HURDA LİSTESİ TEMİZLENDİ."+"\n")                  
    logsclose()
     
    export3("local")
############################################################

############################################################
# Cancelling filter option to return normal sized table.
def returnTable3():
    global p3
    p3 = 0
    
    for i in table3.get_children():
        table3.delete(i)

    
    for i in range(0,len(new_table3)):
        table3.insert(parent = '', index=index3[i],id=i+1, values=[new_table3[i][j] for j in range(0,len(new_table3[0]))])

           
    filter_button3['state'] = ACTIVE
    filter_cancel3['state'] = DISABLED  
    excel_button3['state'] = ACTIVE


    export3("local") 
    updateTable(table3,"Hurda Listesi.xlsx")

def returnTable2():
    global p2
    p2=0
    extratable = []
    
    for i in table2.get_children():
        extratable.append(table2.item(i).get("values"))
        

    for i in table2.get_children():
        table2.delete(i)

    newlist = filtertable + extratable
    
    for i in range(0,len(newlist)):
        table2.insert(parent = '', index='end',id=i+1, values=[newlist[i][j] for j in range(0,len(newlist[0]))])

    
    sorting(table2)
        
    filter_button2['state'] = ACTIVE
    filter_cancel2['state'] = DISABLED  
    excel_button2['state'] = ACTIVE

    export2("local") 
    updateTable(table2,"Zimmet Listesi.xlsx")

def returnTable():

    global p
    p = 0

    extratable = []
    
    for i in table.get_children():
        extratable.append(table.item(i).get("values"))
        

    for i in table.get_children():
        table.delete(i)

    newlist = new_table + extratable
    
    for i in range(0,len(newlist)):
        table.insert(parent = '', index='end',id=i+1, values=[newlist[i][j] for j in range(0,len(newlist[0]))])

    
    sorting(table)
        
    filter_button['state'] = ACTIVE
    filter_cancel['state'] = DISABLED  
    excel_button['state'] = ACTIVE

    export("local") 
    updateTable(table,"Depo Listesi.xlsx")

############################################################

############################################################
# Filtering products by their name or person.
def filter3():
    global p3
    p3 = 0
    flag = False
    if category_entry3.get().isspace() or category_entry3.get() == '':
        messagebox.showwarning("UYARI","Lütfen Ürün İsmi Girin!")
        flag = True
    
    global new_table3,index3

    new_table3 = []
    index3 = []

    if flag == False:
        p3 = 1
        for i in table3.get_children():
            index3.append(table3.index(i))
    
        for i in table3.get_children():
            
            if str(entryUpper(category_entry3.get())) not in str(entryUpper(table3.item(i).get("values")[1])):
                new_table3.append(table3.item(i).get("values"))
                table3.delete(i)

            else:
                new_table3.append(table3.item(i).get("values"))
                
            
            
        filter_button3['state'] = DISABLED
        filter_cancel3['state'] = ACTIVE
        excel_button3['state'] = DISABLED



def filter2():
    global filtertable
    global p2
    p2 = 0
    if control_menu.get() == "Seç:":
        messagebox.showwarning("UYARI","Lütfen Kategori Seçiniz!")

    elif control_menu.get() == "Ürün: ":
        ctrl = False
        
        filtertable = []

        if category_entry2.get().isspace() or category_entry2.get() == '':
            messagebox.showwarning("UYARI","Lütfen Ürün İsmi Girin!")
            ctrl = True

        if ctrl == False:
            p2 = 1    
            for i in table2.get_children():
                if str(entryUpper(category_entry2.get())) not in str(entryUpper(table2.item(i).get("values")[1])):
                    filtertable.append(table2.item(i).get("values"))
                
                    table2.delete(i)

            filter_button2['state'] = DISABLED
            filter_cancel2['state'] = ACTIVE
            excel_button2['state'] = DISABLED

    elif control_menu.get() == "Kişi: ":
        ctrl = False
        filtertable = []

        if category_entry2.get().isspace() or category_entry2.get() == '':
            messagebox.showwarning("UYARI","Lütfen Kişi İsmi Girin!")
            ctrl = True

        if ctrl == False:
            p2= 1   
            for i in table2.get_children():
                if str(entryUpper(category_entry2.get())) not in str(entryUpper(table2.item(i).get("values")[4])) and str(entryUpper(category_entry2.get())) not in str(entryUpper(table2.item(i).get("values")[5])):
                    filtertable.append(table2.item(i).get("values"))
                
                    table2.delete(i)

            filter_button2['state'] = DISABLED
            filter_cancel2['state'] = ACTIVE
            excel_button2['state'] = DISABLED



def filter():

    global p
    p = 0
    
    flag = False
    
    if category_entry.get().isspace() or category_entry.get() == '':
        messagebox.showwarning("UYARI","Lütfen Ürün İsmi Girin!")
        flag = True
    
    global new_table

    new_table = []

    if flag == False:
        p = 1
        for i in table.get_children():
            if str(entryUpper(category_entry.get())) not in str(entryUpper(table.item(i).get("values")[1])):
                new_table.append(table.item(i).get("values"))
                
                table.delete(i)
                        
        filter_button['state'] = DISABLED
        filter_cancel['state'] = ACTIVE
        excel_button['state'] = DISABLED

    #print(entryUpper(category_entry))

############################################################

############################################################
# Prevents entring nonnumeric values to amount of product entries.
def testVal(inStr,acttyp):
    if acttyp == '1': #insert
        if not inStr.isdigit():
            return False
    return True
############################################################

############################################################
# Adding product back to inventory from registered list by clicking with mouse.
def addBack2():
    flag = False
    flag2 = False
    try:
        iselected = table2.focus()

        for o in table.get_children():         
            if(str(table.item(o).get("values")[0]) == str(table2.item(iselected).get("values")[0]) 
            and str(table.item(o).get("values")[1]) == str(table2.item(iselected).get("values")[1])
            and str(table.item(o).get("values")[3]) == str(table2.item(iselected).get("values")[3]) 
            and str(table.item(o).get("values")[4]) == aciklama02.get()):
                flag2 = True
                break



        if miktar02.get().isspace() or miktar02.get() == '':
            messagebox.showwarning("UYARI!","Lütfen Miktar Girin!")
            flag = True
            
        else:
            if (str(table2.item(iselected).get("values")[0]) == str(malzeme0x.cget("text")) 
            and str(table2.item(iselected).get("values")[1]) == str(malzeme_metin0x.cget("text"))
            and str(table2.item(iselected).get("values")[3]) == str(olcu0x.cget("text"))
            and str(table2.item(iselected).get("values")[4]) == str(veren0x.cget("text"))
            and str(table2.item(iselected).get("values")[5]) == str(alan0x.cget("text"))
            and str(table2.item(iselected).get("values")[6]) == str(tarih0x.cget("text"))):

                if flag2 == True:
                    if int(miktar02.get()) >= int(table2.item(iselected).get("values")[2]):
                        table.item(o,values=(table.item(o).get("values")[0],table.item(o).get("values")[1],int(table.item(o).get("values")[2])
                        +int(table2.item(iselected).get("values")[2]),table.item(o).get("values")[3],table.item(o).get("values")[4]))

                        logs()
                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(iselected).get("values")[0])+"  -  "+str(table2.item(iselected).get("values")[1])+"  -  "
                        +str(table2.item(iselected).get("values")[2])+"  -  "+str(table2.item(iselected).get("values")[3])+"  -  "+str(table2.item(iselected).get("values")[4])
                        +"  -  "+str(table2.item(iselected).get("values")[5])+"  -  "+str(table2.item(iselected).get("values")[6])+"  -  "+str(table2.item(iselected).get("values")[7])+"\n")                  
                        logsclose()


                    else:
                        table.item(o,values=(table.item(o).get("values")[0],table.item(o).get("values")[1],int(table.item(o).get("values")[2])
                        +int(miktar02.get()),table.item(o).get("values")[3],table.item(o).get("values")[4]))

                        logs()
                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(iselected).get("values")[0])+"  -  "+str(table2.item(iselected).get("values")[1])+"  -  "
                        +miktar02.get()+"  -  "+str(table2.item(iselected).get("values")[3])+"  -  "+str(table2.item(iselected).get("values")[4])
                        +"  -  "+str(table2.item(iselected).get("values")[5])+"  -  "+str(table2.item(iselected).get("values")[6])+"  -  "+str(table2.item(iselected).get("values")[7])+"\n")                  
                        logsclose()

                else:
                
                    if int(miktar02.get()) >= int(table2.item(iselected).get("values")[2]):
                        try:

                            table.insert(parent = '', index="end",id=max([int(q) for q in table.get_children()])+1,values=(str(table2.item(iselected).get("values")[0]),
                            table2.item(iselected).get("values")[1],int(table2.item(iselected).get("values")[2]),table2.item(iselected).get("values")[3],aciklama02.get()))
                            sorting(table)
                        except:
                            table.insert(parent = '', index="end",id=1,values=(table2.item(iselected).get("values")[0],table2.item(iselected).get("values")[1]
                            ,int(table2.item(iselected).get("values")[2]),table2.item(iselected).get("values")[3],aciklama02.get()))
                            sorting(table)
                        
                        logs()
                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(iselected).get("values")[0])+"  -  "+str(table2.item(iselected).get("values")[1])+"  -  "
                        +str(table2.item(iselected).get("values")[2])+"  -  "+str(table2.item(iselected).get("values")[3])+"  -  "+str(table2.item(iselected).get("values")[4])
                        +"  -  "+str(table2.item(iselected).get("values")[5])+"  -  "+str(table2.item(iselected).get("values")[6])+"  -  "+aciklama02.get()+"\n")                  
                        logsclose()

                    else:
                        try:
                            table.insert(parent = '', index="end",id=1,values=(table2.item(iselected).get("values")[0],table2.item(iselected).get("values")[1]
                                    ,miktar02.get(),table2.item(iselected).get("values")[3],aciklama02.get()))
                            sorting(table)
                        except:
                            table.insert(parent = '', index="end",id=max([int(q) for q in table.get_children()])+1,values=(str(table2.item(iselected).get("values")[0]),
                                    table2.item(iselected).get("values")[1],miktar02.get(),table2.item(iselected).get("values")[3],aciklama02.get()))
                            sorting(table)

                        logs()
                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(iselected).get("values")[0])+"  -  "+str(table2.item(iselected).get("values")[1])+"  -  "
                        +miktar02.get()+"  -  "+str(table2.item(iselected).get("values")[3])+"  -  "+str(table2.item(iselected).get("values")[4])
                        +"  -  "+str(table2.item(iselected).get("values")[5])+"  -  "+str(table2.item(iselected).get("values")[6])+"  -  "+aciklama02.get()+"\n")                  
                        logsclose()
                

                if(int(miktar02.get())) >= int(table2.item(int(iselected)).get("values")[2]):
                    real_table3 = []
                    indexlist = []
                    table2.delete(iselected)

                    for i in table2.get_children():
                        real_table3.append(table2.item(i).get("values"))
                        indexlist.append(i)


                    for x in table2.get_children():
                        table2.delete(x)

                    
                    for l in range(0,len(real_table3)):    
                        table2.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])
                        
                    sorting(table2) 

                else:
                    table2.item(iselected, text="",values=(table2.item(iselected).get("values")[0],table2.item(iselected).get("values")[1],
                    table2.item(iselected).get("values")[2]-int(miktar02.get())
                    ,table2.item(iselected).get("values")[3],table2.item(iselected).get("values")[4],table2.item(iselected).get("values")[5],
                    table2.item(iselected).get("values")[6],table2.item(iselected).get("values")[7]))   
            
                export("local")
                export2("local")
            else:
                messagebox.showwarning("UYARI","Farklı Bir Ürün Seçili!")  
            
           
    except IndexError:
        messagebox.showwarning("UYARI","Ürün Seçili Değil!")
############################################################


############################################################
# Adding product back to inventory from registered list by its product number.
def addBack():
    try:
        global real_table6
        real_table6 = []

        flag = False

        flag2 = False

        
        
        for j in table2.get_children():

            if(malzeme0.get() == str(table2.item(j).get("values")[0]) and veren0.get() == str(table2.item(j).get("values")[4])
            and alan0.get() == str(table2.item(j).get("values")[5])):
                flag = True
                break

        if flag == True:
            for i in table2.get_children():
                
                    if(malzeme0.get().isspace() or malzeme0.get() == '' or miktar0.get().isspace() or miktar0.get() == ''
                    or veren0.get().isspace() or veren0.get() == '' or alan0.get().isspace() or alan0.get() == ''):
                        
                        messagebox.showwarning("UYARI","Parametreleri lütfen doldurun!")
                        break
                    
                    elif(str(table2.item(i).get("values")[0]) == malzeme0.get()):
                        

                        for o in table.get_children():         
                            if(str(table.item(o).get("values")[0]) == str(table2.item(i).get("values")[0]) 
                            and str(table.item(o).get("values")[1]) == str(table2.item(i).get("values")[1])
                            and str(table.item(o).get("values")[3]) == str(table2.item(i).get("values")[3]) 
                            and str(table.item(o).get("values")[4]) == aciklama0.get()):
                                flag2 = True
                                break

                        

                               
                        if(table2.item(i).get("values")[2] - int(miktar0.get())) <=0:
                            if flag2 == True:
                                if int(miktar0.get()) >= int(table2.item(i).get("values")[2]):
                                    table.item(o,values=(table.item(o).get("values")[0],table.item(o).get("values")[1],int(table.item(o).get("values")[2])
                                    +int(table2.item(i).get("values")[2]),table.item(o).get("values")[3],table.item(o).get("values")[4]))

                                    logs()
                                    logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                    +str(table2.item(i).get("values")[2])+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                    +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+aciklama0.get()+"\n")                  
                                    logsclose()


                                else:
                                    table.item(o,values=(table.item(o).get("values")[0],table.item(o).get("values")[1],int(table.item(o).get("values")[2])
                                    +int(miktar0.get()),table.item(o).get("values")[3],table.item(o).get("values")[4]))

                                    logs()
                                    logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                    +miktar0.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                    +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+aciklama0.get()+"\n")                  
                                    logsclose()


                            else:
                                if int(miktar0.get()) >= int(table2.item(i).get("values")[2]):
                                    if aciklama0.get().isspace() or aciklama0.get() == '':
                                        try:    
                                            table.insert(parent='',index = 'end',id = max([int(q) for q in table.get_children()])+1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(table2.item(i).get("values")[2]),table2.item(i).get("values")[3],table2.item(i).get("values")[7]))
                                        except:
                                            table.insert(parent='',index = 'end',id = 1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(table2.item(i).get("values")[2]),table2.item(i).get("values")[3],table2.item(i).get("values")[7]))

                                        logs()
                                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                        +str(table2.item(i).get("values")[2])+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                        +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+str(table2.item(i).get("values")[7])+"\n")                  
                                        logsclose()
                                                    
                                    else:
                                        try:    
                                            table.insert(parent='',index = 'end',id = max([int(q) for q in table.get_children()])+1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(table2.item(i).get("values")[2]),table2.item(i).get("values")[3],aciklama0.get()))
                                        except:
                                            table.insert(parent='',index = 'end',id = 1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(table2.item(i).get("values")[2]),table2.item(i).get("values")[3],aciklama0.get()))

                                        logs()
                                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                        +str(table2.item(i).get("values")[2])+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                        +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+aciklama0.get()+"\n")                  
                                        logsclose()

                                else:
                                    if aciklama0.get().isspace() or aciklama0.get() == '':
                                        try:    
                                            table.insert(parent='',index = 'end',id = max([int(q) for q in table.get_children()])+1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],table2.item(i).get("values")[7]))
                                        except:
                                            table.insert(parent='',index = 'end',id = 1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],table2.item(i).get("values")[7]))

                                        logs()
                                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                        +miktar0.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                        +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+str(table2.item(i).get("values")[7])+"\n")                  
                                        logsclose()
                                                    
                                    else:
                                        try:   
                                            table.insert(parent='',index = 'end',id = max([int(q) for q in table.get_children()])+1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],aciklama0.get()))
                                        except:
                                            table.insert(parent='',index = 'end',id = 1,values=(table2.item(i).get("values")[0],
                                            table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],aciklama0.get()))

                                        logs()
                                        logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                        +miktar0.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                        +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+aciklama0.get()+"\n")                  
                                        logsclose()



                            real_table3 = []
                            indexlist = []
                            table2.delete(i)

                            for c in table2.get_children():
                                real_table3.append(table2.item(c).get("values"))
                                indexlist.append(c)


                            for x in table2.get_children():
                                table2.delete(x)

                                
                            for l in range(0,len(real_table3)):    
                                table2.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])

                                
                                
                            sorting(table)
                            sorting(table2)
                            break
                        else:
                                ## insert kısmı
                            if flag2 == True:
                                table.item(o,values=(table.item(o).get("values")[0],table.item(o).get("values")[1],int(table.item(o).get("values")[2])
                                +int(miktar0.get()),table.item(o).get("values")[3],table.item(o).get("values")[4]))

                                logs()
                                logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                +miktar0.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+aciklama0.get()+"\n")                  
                                logsclose()


                                
                            else:
                                if aciklama0.get().isspace() or aciklama0.get() == '':
                                    try:    
                                        table.insert(parent='',index = 'end',id = max([int(q) for q in table.get_children()])+1,values=(table2.item(i).get("values")[0],
                                        table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],table2.item(i).get("values")[7]))
                                    except:
                                        table.insert(parent='',index = 'end',id = 1,values=(table2.item(i).get("values")[0],
                                        table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],table2.item(i).get("values")[7]))

                                    logs()
                                    logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                    +miktar0.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                    +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+str(table2.item(i).get("values")[7])+"\n")                  
                                    logsclose()
                                                
                                else:
                                    try:    
                                        table.insert(parent='',index = 'end',id = max([int(q) for q in table.get_children()])+1,values=(table2.item(i).get("values")[0],
                                        table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],aciklama0.get()))
                                    except:
                                        table.insert(parent='',index = 'end',id = 1,values=(table2.item(i).get("values")[0],
                                        table2.item(i).get("values")[1],int(miktar0.get()),table2.item(i).get("values")[3],aciklama0.get()))

                                    logs()
                                    logfile.write("ZİMMET LİSTESİ"+"\t"+"GERİ EKLEME"+"    "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                    +miktar0.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                    +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+aciklama0.get()+"\n")                  
                                    logsclose()

                                ##
                            table2.item(i,values=(table2.item(i).get("values")[0],table2.item(i).get("values")[1],table2.item(i).get("values")[2]
                            -int(miktar0.get()),table2.item(i).get("values")[3],table2.item(i).get("values")[4],table2.item(i).get("values")[5]
                            ,table2.item(i).get("values")[6],table2.item(i).get("values")[7]))
                            #table.item(i).get("values")[2] +=1

                            sorting(table)
                            sorting(table2)

                            
                                
                            break
                        break

            export("local")
            export2("local")

        elif flag == False:
            messagebox.showwarning("UYARI","Böyle Bir Ürün Yok!")
    except IndexError:
        messagebox.showwarning("UYARI","Böyle Bir Ürün Yok!")
    except ValueError:
        messagebox.showwarning("UYARI","Böyle Bir Ürün Yok!")

############################################################    

############################################################
# Sorting products according to their product id on the tables.
def sorting(table_list):
    
    if p != 1 and p2 !=1:
        try:
            sort_list = [table_list.item(a).get("values") for a in table_list.get_children()]

            intlist = []
            strlist = []
            for i in range(0,len(sort_list)):
                if type(sort_list[i][0]) == int:
                    intlist.append(int(sort_list[i][0]))
                else:
                    strlist.append(sort_list[i][0])  
                

            intlist.sort()
            
            real_table4 = []
            sortlist2 = []

            for i in range(0,len(intlist)):
                for j in table_list.get_children():
                    if str(intlist[i]) == str(table_list.item(j).get("values")[0]):
                        sortlist2.append(table_list.item(j).get("values"))
                        table_list.delete(j)
                        break            
                        
            
            for i in range(0,len(strlist)):
                for j in table_list.get_children():
                    if str(strlist[i]) == str(table_list.item(j).get("values")[0]):
                        sortlist2.append(table_list.item(j).get("values"))
                        table_list.delete(j)
                        break

            
            for x in table_list.get_children():
                table_list.delete(x)

            for i in range(0,len(sortlist2)):
                table_list.insert(parent = '', index='end',id=i+1, values=[sortlist2[i][j] for j in range(0,len(sortlist2[0]))])
        except:
            print("Sorting Sorunu")
    coloring()
    coloring2()
    coloring3()

############################################################

############################################################
# Editting mouse-selected registered product.
def edit4():

    logs()
    logfile.write("ZİMMET LİSTESİ"+"\t"+"DÜZENLEME"+"      "+str(table2.item(selectedItemx).get("values")[0])+"  -  "+str(table2.item(selectedItemx).get("values")[1])+"  -  "
    +str(table2.item(selectedItemx).get("values")[2])+"  -  "+str(table2.item(selectedItemx).get("values")[3])+"  -  "+str(table2.item(selectedItemx).get("values")[4]
    +"  -  "+str(table2.item(selectedItemx).get("values")[5])+"  -  "+str(table2.item(selectedItemx).get("values")[6])+"  -  "+str(table2.item(selectedItemx).get("values")[7]))) 


    table2.item(selectedItemx, text="",values=(malzemet.get(),malzeme_metint.get()
            ,miktart.get(),olcut.get(),verent.get(),alant.get(),tariht.get(),aciklamat.get()))

    logfile.write("   -------YENİ ÜRÜN:-------   "+malzemet.get()+"  -  "+malzeme_metint.get()+"  -  "
    +miktart.get()+"  -  "+olcut.get()+"  -  "+verent.get()+"  -  "+alant.get()+"  -  "+tariht.get()+"  -  "+aciklamat.get()+"\n")                  
    logsclose()
    
    sorting(table2)

    export2("local")
############################################################

############################################################
# Selecting product with mouse to edit it.
def selectReg():
    try:
        global selectedItemx
        selectedItemx = table2.focus()
        

        malzemet.delete(0,END)
        malzeme_metint.delete(0,END)
        miktart.delete(0,END)
        olcut.delete(0,END)
        verent.delete(0,END)
        alant.delete(0,END)
        tariht.delete(0,END)
        aciklamat.delete(0,END)

        malzemet.insert(0,str(table2.item(selectedItemx).get("values")[0]))
        malzeme_metint.insert(0,str(table2.item(selectedItemx).get("values")[1]))
        miktart.insert(0,str(table2.item(selectedItemx).get("values")[2]))
        olcut.insert(0,str(table2.item(selectedItemx).get("values")[3]))
        verent.insert(0,str(table2.item(selectedItemx).get("values")[4]))
        alant.insert(0,str(table2.item(selectedItemx).get("values")[5]))
        tariht.insert(0,str(table2.item(selectedItemx).get("values")[6]))
        aciklamat.insert(0,str(table2.item(selectedItemx).get("values")[7]))

        duzenlet['state'] = ACTIVE

    except IndexError:
        messagebox.showwarning("UYARI","Lütfen Ürün Seçin!")
############################################################

############################################################
# Editting registered product which searched by its product number.
def edit3():

    logs()
    logfile.write("ZİMMET LİSTESİ"+"\t"+"DÜZENLEME"+"      "+str(table2.item(productR).get("values")[0])+"  -  "+str(table2.item(productR).get("values")[1])+"  -  "
    +str(table2.item(productR).get("values")[2])+"  -  "+str(table2.item(productR).get("values")[3])+"  -  "+str(table2.item(productR).get("values")[4]
    +"  -  "+str(table2.item(productR).get("values")[5])+"  -  "+str(table2.item(productR).get("values")[6])+"  -  "+str(table2.item(productR).get("values")[7]))) 


    table2.item(productR, text="",values=(malzemeq2.get(),malzeme_metinq.get()
            ,miktarq.get(),olcuq.get(),verenq.get(),alanq.get(),tarihq.get(),aciklamaq.get()))

    logfile.write("   -------YENİ ÜRÜN:-------   "+malzemeq2.get()+"  -  "+malzeme_metinq.get()+"  -  "
    +miktarq.get()+"  -  "+olcuq.get()+"  -  "+verenq.get()+"  -  "+alanq.get()+"  -  "+tarihq.get()+"  -  "+aciklamaq.get()+"\n")                  
    logsclose()


    sorting(table2)

    export2("local")

############################################################
# Searching product by its product id.
def searchRegist():
    try:
        ctrl = False
        global productR
        productR = 0
        for i in table2.get_children():
            if(malzemeq.get() == str(table2.item(i).get("values")[0])):
                productR = i
                malzemeq['state'] = DISABLED

                malzemeq2.delete(0,END)
                malzeme_metinq.delete(0,END)
                miktarq.delete(0,END)
                olcuq.delete(0,END)
                aciklamaq.delete(0,END)
                verenq.delete(0,END)
                alanq.delete(0,END)
                tarihq.delete(0,END)

                duzenleq['state'] = ACTIVE

                malzemeq2.insert(0,str(table2.item(i).get("values")[0]))
                malzeme_metinq.insert(0,str(table2.item(i).get("values")[1]))
                miktarq.insert(0,str(table2.item(i).get("values")[2]))
                olcuq.insert(0,str(table2.item(i).get("values")[3]))
                aciklamaq.insert(0,str(table2.item(i).get("values")[7]))
                verenq.insert(0,str(table2.item(i).get("values")[4]))
                alanq.insert(0,str(table2.item(i).get("values")[5]))
                tarihq.insert(0,str(table2.item(i).get("values")[6]))
                ctrl = True

                break
        
        if ctrl == False:
            messagebox.showwarning("UYARI","Ürün Mevcut Değil!")
            
    except IndexError:
        messagebox.showwarning("UYARI","Ürün Bulunamadı!")
    except ValueError:
        messagebox.showwarning("UYARI","Ürün Bulunamadı!")
############################################################

############################################################
# Deleting mouse-selected registered product on the table.
def deleteSelectedReg():
    
    try:
        
        selected = table2.focus()
        
        if miktarr2.get().isspace() or miktarr2.get() == '':
            if table2.item(selected).get("values")[2]<=1:
                try:
                    table3.insert(parent = '', index=0,id=1,values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,1,table2.item(selected).get("values")[3],str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))

                    
                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,
                    values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,1,table2.item(selected).get("values")[3],
                    str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))
                
                logs()
                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(selected).get("values")[0])+"  -  "+str(table2.item(selected).get("values")[1])+"  -  "
                +"1"+"  -  "+str(table2.item(selected).get("values")[3])+"  -  "+str(table2.item(selected).get("values")[4])
                +"  -  "+str(table2.item(selected).get("values")[5])+"  -  "+str(table2.item(selected).get("values")[6])+"  -  "+hurdaaciklama4.get()+"\n")                  
                logsclose()
                
                real_table3 = []
                indexlist = []
                table2.delete(selected)

                for i in table2.get_children():
                    real_table3.append(table2.item(i).get("values"))
                    indexlist.append(i)


                for x in table2.get_children():
                    table2.delete(x)

                
                for l in range(0,len(real_table3)):    
                    table2.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])
                    
                
            else:

                try:
                    table3.insert(parent = '', index=0,id=1,values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,1,table2.item(selected).get("values")[3],str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))
                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,
                    values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,1,table2.item(selected).get("values")[3],
                    str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))

                logs()
                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(selected).get("values")[0])+"  -  "+str(table2.item(selected).get("values")[1])+"  -  "
                +"1"+"  -  "+str(table2.item(selected).get("values")[3])+"  -  "+str(table2.item(selected).get("values")[4])
                +"  -  "+str(table2.item(selected).get("values")[5])+"  -  "+str(table2.item(selected).get("values")[6])+"  -  "+hurdaaciklama4.get()+"\n")                  
                logsclose()

                    
                table2.item(selected, text="",values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1],table2.item(selected).get("values")[2]-1
                ,table2.item(selected).get("values")[3],table2.item(selected).get("values")[4],table2.item(selected).get("values")[5]
                            ,table2.item(selected).get("values")[6],table2.item(selected).get("values")[7]))       
        else:
            if(table2.item(selected).get("values")[2] - int(miktarr2.get())) <=0:
                try:
                    table3.insert(parent = '', index=0,id=1,values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,int(table2.item(selected).get("values")[2]),table2.item(selected).get("values")[3],str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))
                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,
                    values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,int(table2.item(selected).get("values")[2]),table2.item(selected).get("values")[3],
                    str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))

                logs()
                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(selected).get("values")[0])+"  -  "+str(table2.item(selected).get("values")[1])+"  -  "
                +str(table2.item(selected).get("values")[2])+"  -  "+str(table2.item(selected).get("values")[3])+"  -  "+str(table2.item(selected).get("values")[4])
                +"  -  "+str(table2.item(selected).get("values")[5])+"  -  "+str(table2.item(selected).get("values")[6])+"  -  "+hurdaaciklama4.get()+"\n")                  
                logsclose()

                real_table4 = []
                indexlist2 = []
                table2.delete(selected)

                for i in table2.get_children():
                    real_table4.append(table2.item(i).get("values"))
                    indexlist2.append(i)


                for x in table2.get_children():
                    table2.delete(x)

                
                for l in range(0,len(real_table4)):    
                    table2.insert(parent = '', index='end',id=indexlist2[l],values=[real_table4[l][a] for a in range (0,len(real_table4[0]))])

            else:

                try:
                    table3.insert(parent = '', index=0,id=1,values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,int(miktarr2.get()),table2.item(selected).get("values")[3],str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))
                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,
                    values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                    ,int(miktarr2.get()),table2.item(selected).get("values")[3],
                    str(str(hurdaaciklama4.get())+"     "+str(datetime.now())[0:19])))

                logs()
                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(selected).get("values")[0])+"  -  "+str(table2.item(selected).get("values")[1])+"  -  "
                +miktarr2.get()+"  -  "+str(table2.item(selected).get("values")[3])+"  -  "+str(table2.item(selected).get("values")[4])
                +"  -  "+str(table2.item(selected).get("values")[5])+"  -  "+str(table2.item(selected).get("values")[6])+"  -  "+hurdaaciklama4.get()+"\n")                  
                logsclose()
                

                table2.item(selected, text="",values=(table2.item(selected).get("values")[0],table2.item(selected).get("values")[1]
                ,table2.item(selected).get("values")[2]-int(miktarr2.get())
                ,table2.item(selected).get("values")[3],table2.item(selected).get("values")[4],table2.item(selected).get("values")[5]
                            ,table2.item(selected).get("values")[6],table2.item(selected).get("values")[7]))
        sorting(table2)
        
        export2("local")
        export3("local")

    except IndexError:
        messagebox.showwarning("UYARI","Lütfen Ürün Seçin!")
############################################################

############################################################
# Deleting registered product which searched by its product id.
def deleteReg():
    try:
        global real_table5
        real_table5 = []

        flag = False
        
        for j in table2.get_children():

            if(malzemer.get() == str(table2.item(j).get("values")[0])):
                flag = True
                break

        if flag == True:
            for i in table2.get_children():
                
                    if(malzemer.get().isspace() or malzemer.get() == '' or miktarr.get().isspace() or miktarr.get() == ''):
                        
                        messagebox.showwarning("UYARI","Parametreyi lütfen doldurun!")
                        break
                    
                    elif(str(table2.item(i).get("values")[0]) == malzemer.get()):
                    
                        if(table2.item(i).get("values")[2] - int(miktarr.get())) <=0:

                            try:
                                table3.insert(parent = '', index=0,id=1,values=(table2.item(i).get("values")[0],table2.item(i).get("values")[1]
                                ,int(table2.item(i).get("values")[2]),table2.item(i).get("values")[3],str(str(hurdaaciklama3.get())+"     "+str(datetime.now())[0:19])))


                                logs()
                                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                +str(table2.item(i).get("values")[2])+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+hurdaaciklama3.get()+"\n")                  
                                logsclose()


                            except:
                                table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table2.item(i).get("values")[0],table2.item(i).get("values")[1]
                                ,int(table2.item(i).get("values")[2]),table2.item(i).get("values")[3],str(str(hurdaaciklama3.get())+"     "+str(datetime.now())[0:19])))

                                logs()
                                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                +str(table2.item(i).get("values")[2])+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+hurdaaciklama3.get()+"\n")                  
                                logsclose()

                            real_table3 = []
                            indexlist= []
                            table2.delete(i)

                            for j in table2.get_children():
                                real_table3.append(table2.item(j).get("values"))
                                indexlist.append(j)
                                
                            
                            for x in table2.get_children():    
                                table2.delete(x)
                                
                            
                            for l in range(0,len(real_table3)):    
                                table2.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])
                                
                            break
                        else:
                            try:
                                table3.insert(parent = '', index=0,id=1,values=(table2.item(i).get("values")[0],table2.item(i).get("values")[1]
                                ,int(miktarr.get()),table2.item(i).get("values")[3],str(str(hurdaaciklama3.get())+"     "+str(datetime.now())[0:19])))

                                logs()
                                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                +miktarr.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+hurdaaciklama3.get()+"\n")                  
                                logsclose()


                            except:
                                table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table2.item(i).get("values")[0],table2.item(i).get("values")[1]
                                ,int(int(miktarr.get())),table2.item(i).get("values")[3],str(str(hurdaaciklama3.get())+"     "+str(datetime.now())[0:19]))) 

                                logs()
                                logfile.write("ZİMMET LİSTESİ"+"\t"+"SİLME"+"          "+str(table2.item(i).get("values")[0])+"  -  "+str(table2.item(i).get("values")[1])+"  -  "
                                +miktarr.get()+"  -  "+str(table2.item(i).get("values")[3])+"  -  "+str(table2.item(i).get("values")[4])
                                +"  -  "+str(table2.item(i).get("values")[5])+"  -  "+str(table2.item(i).get("values")[6])+"  -  "+hurdaaciklama3.get()+"\n")                  
                                logsclose()


                            table2.item(i,values=(table2.item(i).get("values")[0],table2.item(i).get("values")[1],table2.item(i).get("values")[2]
                            -int(miktarr.get()),table2.item(i).get("values")[3],table2.item(i).get("values")[4],table2.item(i).get("values")[5]
                            ,table2.item(i).get("values")[6],table2.item(i).get("values")[7]))
                            #table.item(i).get("values")[2] +=1
                        break
            sorting(table2)

            export2("local")
            export3("local")                    
        
        elif flag == False:
            messagebox.showwarning("UYARI","Böyle Bir Ürün Yok!")
    except IndexError:
        messagebox.showwarning("UYARI","Böyle Bir Ürün Yok!")
    except ValueError:
        messagebox.showwarning("UYARI","Böyle Bir Ürün Yok!")
############################################################

############################################################
# Cleaning entries and back buttons on the interface.
def clearReg():
    malzemer.delete(0,END)
    miktarr.delete(0,END)
    miktarr2.delete(0,END)
    hurdaaciklama3.delete(0,END)
    hurdaaciklama4.delete(0,END)

def backReg():
    remove_register_frame.place_forget()
    malzemer.delete(0,END)
    miktarr.delete(0,END)
    miktarr2.delete(0,END)
    hurdaaciklama3.delete(0,END)
    hurdaaciklama4.delete(0,END)

def clearAll():
    malzeme0.delete(0,END)
    miktar0.delete(0,END)
    aciklama0.delete(0,END)
    veren0.delete(0,END)
    alan0.delete(0,END)
    miktar02.delete(0,END)
    aciklama02.delete(0,END)

    malzeme0x.configure(text="")
    malzeme_metin0x.configure(text="")
    olcu0x.configure(text="")
    veren0x.configure(text="")
    alan0x.configure(text="")
    tarih0x.configure(text="")

    ekle02['state'] = DISABLED
def backMain():

    malzeme0.delete(0,END)
    miktar0.delete(0,END)
    aciklama0.delete(0,END)
    veren0.delete(0,END)
    alan0.delete(0,END)
    miktar02.delete(0,END)
    aciklama02.delete(0,END)

    malzeme0x.configure(text="")
    malzeme_metin0x.configure(text="")
    olcu0x.configure(text="")
    veren0x.configure(text="")
    alan0x.configure(text="")
    tarih0x.configure(text="")
    ekle02['state'] = DISABLED
    back_register_frame.place_forget()

############################################################

############################################################
# Selecting product on registered table.
def selecting():
    try:
        selectinq = table2.focus()

        miktar02.delete(0,END)
        aciklama02.delete(0,END)

        malzeme0x.configure(text=table2.item(selectinq).get("values")[0],fg="red")
        malzeme_metin0x.configure(text=table2.item(selectinq).get("values")[1],fg="red",font=("",7))
        miktar02.insert(0,table2.item(selectinq).get("values")[2])
        olcu0x.configure(text=table2.item(selectinq).get("values")[3],fg="red")
        veren0x.configure(text=table2.item(selectinq).get("values")[4],fg="red")
        alan0x.configure(text=table2.item(selectinq).get("values")[5],fg="red")
        tarih0x.configure(text=table2.item(selectinq).get("values")[6],fg="red")
        aciklama02.insert(0,table2.item(selectinq).get("values")[7])
        ekle02['state'] = ACTIVE

    except IndexError:
        messagebox.showerror("UYARI","Ürün Seçilmedi!")
############################################################

############################################################
# Interface of  adding registered product back to inventory.
def backInventory():
    back_register_frame.place(x=0,y=0)

    malzeme_label = Label(back_register_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=16)
    malzeme_label = Label(back_register_frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=39)
    malzeme_label = Label(back_register_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=64)
    malzeme_label = Label(back_register_frame,text="Teslim Eden: ",bg="LavenderBlush2").place(x=5,y=89)
    malzeme_label = Label(back_register_frame,text="Teslim Alan: ",bg="LavenderBlush2").place(x=5,y=114)
    
    malzeme0.place(x=100,y=16)
    miktar0.place(x=100,y=39)
    aciklama0.place(x=100,y=64)
    veren0.place(x=100,y=89)
    alan0.place(x=100,y=114)
    ekle0.place(x=80,y=134)

    malzeme_label = Label(back_register_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=210)
    malzeme_label = Label(back_register_frame,text="M.M: ",bg="LavenderBlush2").place(x=0,y=235)
    malzeme_label = Label(back_register_frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=260)
    malzeme_label = Label(back_register_frame,text="Ölçü: ",bg="LavenderBlush2").place(x=5,y=285)
    malzeme_label = Label(back_register_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=310)
    malzeme_label = Label(back_register_frame,text="Teslim Eden: ",bg="LavenderBlush2").place(x=5,y=335)
    malzeme_label = Label(back_register_frame,text="Teslim Alan: ",bg="LavenderBlush2").place(x=5,y=360)
    malzeme_label = Label(back_register_frame,text="Teslim Tarihi: ",bg="LavenderBlush2").place(x=5,y=385)

    malzeme_label = Label(back_register_frame,text="________________SEÇEREK EKLE________________",bg="LavenderBlush2",fg = "indian red").place(x=0,y=160)
    malzeme_label = Label(back_register_frame,text="NO İLE EKLE",bg="LavenderBlush2",fg = "indian red").place(x=5,y=0)

    malzeme0x.place(x=100,y=210)
    malzeme_metin0x.place(x=35,y=238)
    olcu0x.place(x=100,y=285)
    veren0x.place(x=100,y=335)
    alan0x.place(x=100,y=360)
    tarih0x.place(x=100,y=385)

    miktar02.place(x=100,y=260)
    aciklama02.place(x=100,y=310)
    
    sec0.place(x=80,y=182)
    ekle02.place(x=12,y=410)
    geri02.place(x=87,y=410)
    temizle02.place(x=162,y=410)
    
    updateTable(table2,"Zimmet Listesi.xlsx")

############################################################

############################################################
# Adding product to registered list from inventory by mouse-click.
def selectRegister():
    flag = False


    flag2 = False

    for o in table2.get_children():         
        if(str(smalzeme.cget("text")) == str(table2.item(o).get("values")[0]) and str(saciklama.get()) == str(table2.item(o).get("values")[7])
        and str(vereny2.get()) == str(table2.item(o).get("values")[4]) and str(alany2.get()) == str(table2.item(o).get("values")[5])
        and str(tarihy2.get()) == str(table2.item(o).get("values")[6])):
            flag2 = True
            break

    
    try:
        iselected = table.focus()
        if miktary2.get().isspace() or miktary2.get() == '':
            messagebox.showwarning("UYARI!","Lütfen Miktar Girin!")
            flag = True
            
        elif vereny2.get().isspace() or vereny2.get() == '' or alany2.get().isspace() or alany2.get() == '' :
            messagebox.showwarning("UYARI!","Lütfen Personel İsimlerini Giriniz!")
            flag = True

        else:
            if (str(table.item(iselected).get("values")[0]) == str(smalzeme.cget("text")) 
            and str(table.item(iselected).get("values")[1]) == str(smalzeme_metin.cget("text"))
            and str(table.item(iselected).get("values")[3]) == str(solcu.cget("text"))):

                if flag2 == True:
                    if int(miktary2.get()) >= int(table.item(iselected).get("values")[2]):
                        table2.item(o,text="",values=(table2.item(o).get("values")[0],table2.item(o).get("values")[1],int(table.item(iselected).get("values")[2])+table2.item(o).get("values")[2],
                        table2.item(o).get("values")[3],table2.item(o).get("values")[4],table2.item(o).get("values")[5],table2.item(o).get("values")[6],
                        table2.item(o).get("values")[7]))

                        logs()
                        logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table2.item(o).get("values")[0])+"  -  "+str(table2.item(o).get("values")[1])+"  -  "
                        +str(table.item(iselected).get("values")[2])+"  -  "+str(table2.item(o).get("values")[3])+"  -  "+str(table2.item(o).get("values")[4])+"  -  "
                        +str(table2.item(o).get("values")[5])+"  -  "+str(table2.item(o).get("values")[6])+"  -  "+str(table2.item(o).get("values")[7])+"\n")                  
                        logsclose()


                    else:
                        table2.item(o,text="",values=(table2.item(o).get("values")[0],table2.item(o).get("values")[1],int(miktary2.get())+table2.item(o).get("values")[2],
                        table2.item(o).get("values")[3],table2.item(o).get("values")[4],table2.item(o).get("values")[5],table2.item(o).get("values")[6],
                        table2.item(o).get("values")[7]))

                        logs()
                        logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table2.item(o).get("values")[0])+"  -  "+str(table2.item(o).get("values")[1])+"  -  "
                        +miktary2.get()+"  -  "+str(table2.item(o).get("values")[3])+"  -  "+str(table2.item(o).get("values")[4])+"  -  "
                        +str(table2.item(o).get("values")[5])+"  -  "+str(table2.item(o).get("values")[6])+"  -  "+str(table2.item(o).get("values")[7])+"\n")                  
                        logsclose()

                        
                else:
                    if int(miktary2.get()) >= int(table.item(iselected).get("values")[2]):
                        try:
                            table2.insert(parent = '', index='end',id=1,values=(table.item(iselected).get("values")[0],table.item(iselected).get("values")[1]
                            ,int(table.item(iselected).get("values")[2]),table.item(iselected).get("values")[3],vereny2.get(),alany2.get()
                            ,str(tarihy2.get()),saciklama.get()))

                            logs()
                            logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(iselected).get("values")[0])+"  -  "+str(table.item(iselected).get("values")[1])+"  -  "
                            +str(table.item(iselected).get("values")[2])+"  -  "+str(table.item(iselected).get("values")[3])+"  -  "+vereny2.get()+"  -  "
                            +alany2.get()+"  -  "+str(tarihy2.get())+"  -  "+saciklama.get()+"\n")                  
                            logsclose()


                        except:
                            table2.insert(parent = '', index='end',id=max([int(q) for q in table2.get_children()])+1,values=(table.item(iselected).get("values")[0],
                                    table.item(iselected).get("values")[1]
                                    ,int(table.item(iselected).get("values")[2]),table.item(iselected).get("values")[3],vereny2.get(),alany2.get()
                                    ,str(tarihy2.get()),saciklama.get()))

                            logs()
                            logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(iselected).get("values")[0])+"  -  "+str(table.item(iselected).get("values")[1])+"  -  "
                            +str(table.item(iselected).get("values")[2])+"  -  "+str(table.item(iselected).get("values")[3])+"  -  "+vereny2.get()+"  -  "
                            +alany2.get()+"  -  "+str(tarihy2.get())+"  -  "+saciklama.get()+"\n")                  
                            logsclose()


                    else:
                        try:
                            table2.insert(parent = '', index='end',id=1,values=(table.item(iselected).get("values")[0],table.item(iselected).get("values")[1]
                                    ,miktary2.get(),table.item(iselected).get("values")[3],vereny2.get(),alany2.get()
                                    ,str(tarihy2.get()),saciklama.get()))

                            logs()
                            logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(iselected).get("values")[0])+"  -  "+str(table.item(iselected).get("values")[1])+"  -  "
                            +miktary2.get()+"  -  "+str(table.item(iselected).get("values")[3])+"  -  "+vereny2.get()+"  -  "
                            +alany2.get()+"  -  "+str(tarihy2.get())+"  -  "+saciklama.get()+"\n")                  
                            logsclose()

                        except:
                            table2.insert(parent = '', index='end',id=max([int(q) for q in table2.get_children()])+1,values=(table.item(iselected).get("values")[0],
                                    table.item(iselected).get("values")[1]
                                    ,miktary2.get(),table.item(iselected).get("values")[3],vereny2.get(),alany2.get()
                                    ,str(tarihy2.get()),saciklama.get()))

                            logs()
                            logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(iselected).get("values")[0])+"  -  "+str(table.item(iselected).get("values")[1])+"  -  "
                            +miktary2.get()+"  -  "+str(table.item(iselected).get("values")[3])+"  -  "+vereny2.get()+"  -  "
                            +alany2.get()+"  -  "+str(tarihy2.get())+"  -  "+saciklama.get()+"\n")                  
                            logsclose()
                    

                if(int(miktary2.get())) >= int(table.item(int(iselected)).get("values")[2]):

                    real_table3 = []
                    indexlist = []
                    table.delete(iselected)

                    for i in table.get_children():
                        real_table3.append(table.item(i).get("values"))
                        indexlist.append(i)


                    for x in table.get_children():
                        table.delete(x)

                        
                    for l in range(0,len(real_table3)):    
                        table.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])

                else:
                    table.item(iselected, text="",values=(table.item(iselected).get("values")[0],table.item(iselected).get("values")[1],
                    table.item(iselected).get("values")[2]-int(miktary2.get())
                    ,table.item(iselected).get("values")[3],table.item(iselected).get("values")[4]))
            else:
                messagebox.showwarning("UYARI","Farklı Bir Ürün Seçili!")       
            sorting(table)
            sorting(table2)

            export("local")
            export2("local")    
    except IndexError:
        messagebox.showwarning("UYARI","Ürün Seçili Değil!")
############################################################
        
############################################################
# Adding product to registered list from inventory by product id.
def noRegister():
    
    flag = False
    flag2 = False

    for o in table2.get_children():         
        if(str(malzemey.get()) == str(table2.item(o).get("values")[0]) and str(aciklamay.get()) == str(table2.item(o).get("values")[7])
        and str(vereny.get()) == str(table2.item(o).get("values")[4]) and str(alany.get()) == str(table2.item(o).get("values")[5])
        and str(tarihy.get()) == str(table2.item(o).get("values")[6])):
            flag2 = True
            break


    for i in table.get_children():
        if miktary.get().isspace() or miktary.get() == '':
            messagebox.showwarning("UYARI!","Lütfen Miktar Girin!")
            flag = True
            break
        elif vereny.get().isspace() or vereny.get() == '' or alany.get().isspace() or alany.get() == '' :
            messagebox.showwarning("UYARI!","Lütfen Personel İsimlerini Giriniz!")
            flag = True
            break

        else:
            if(malzemey.get() == str(table.item(i).get("values")[0])):
                if flag2 == True:
                    if int(miktary.get()) >= int(table.item(i).get("values")[2]):
                        table2.item(o,text="",values=(table2.item(o).get("values")[0],table2.item(o).get("values")[1],int(table.item(i).get("values")[2])+
                        table2.item(o).get("values")[2],
                        table2.item(o).get("values")[3],table2.item(o).get("values")[4],table2.item(o).get("values")[5],table2.item(o).get("values")[6],
                        table2.item(o).get("values")[7]))


                        logs()
                        logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table2.item(o).get("values")[0])+"  -  "+str(table2.item(o).get("values")[1])+"  -  "
                        +str(table.item(i).get("values")[2])+"  -  "+str(table2.item(o).get("values")[3])+"  -  "+str(table2.item(o).get("values")[4])+"  -  "
                        +str(table2.item(o).get("values")[5])+"  -  "+str(table2.item(o).get("values")[6])+"  -  "+str(table2.item(o).get("values")[7])+"\n")                  
                        logsclose()

                    else:
                        table2.item(o,text="",values=(table2.item(o).get("values")[0],table2.item(o).get("values")[1],int(miktary.get())+table2.item(o).get("values")[2],
                        table2.item(o).get("values")[3],table2.item(o).get("values")[4],table2.item(o).get("values")[5],table2.item(o).get("values")[6],
                        table2.item(o).get("values")[7]))


                        logs()
                        logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table2.item(o).get("values")[0])+"  -  "+str(table2.item(o).get("values")[1])+"  -  "
                        +miktary.get()+"  -  "+str(table2.item(o).get("values")[3])+"  -  "+str(table2.item(o).get("values")[4])+"  -  "
                        +str(table2.item(o).get("values")[5])+"  -  "+str(table2.item(o).get("values")[6])+"  -  "+str(table2.item(o).get("values")[7])+"\n")                  
                        logsclose()


                else:
                    if int(miktary.get()) >= int(table.item(i).get("values")[2]):
                        try:
                            if aciklamay.get().isspace() or aciklamay.get() == '': 
                                table2.insert(parent = '', index='end',id=1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,table.item(i).get("values")[2],table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(table.item(i).get("values")[4])))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +str(table.item(i).get("values")[2])+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+str(table.item(i).get("values")[4])+"\n")                  
                                logsclose()


                            else:
                                table2.insert(parent = '', index='end',id=1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,table.item(i).get("values")[2],table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(aciklamay.get())))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +str(table.item(i).get("values")[2])+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+aciklamay.get()+"\n")                  
                                logsclose()


                        except:
                            if aciklamay.get().isspace() or aciklamay.get() == '':
                                table2.insert(parent = '', index='end',id=max([int(q) for q in table2.get_children()])+1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,table.item(i).get("values")[2],table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(table.item(i).get("values")[4])))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +str(table.item(i).get("values")[2])+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+str(table.item(i).get("values")[4])+"\n")                  
                                logsclose()


                            else:
                                table2.insert(parent = '', index='end',id=max([int(q) for q in table2.get_children()])+1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,table.item(i).get("values")[2],table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(aciklamay.get())))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +str(table.item(i).get("values")[2])+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+aciklamay.get()+"\n")                  
                                logsclose()


                    
                    else:
                        try:
                            if aciklamay.get().isspace() or aciklamay.get() == '': 
                                table2.insert(parent = '', index='end',id=1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,miktary.get(),table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(table.item(i).get("values")[4])))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +miktary.get()+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+str(table.item(i).get("values")[4])+"\n")                  
                                logsclose()


                            else:
                                table2.insert(parent = '', index='end',id=1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,miktary.get(),table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(aciklamay.get())))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +miktary.get()+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+aciklamay.get()+"\n")                  
                                logsclose()


                        except:
                            if aciklamay.get().isspace() or aciklamay.get() == '':
                                table2.insert(parent = '', index='end',id=max([int(q) for q in table2.get_children()])+1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,miktary.get(),table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(table.item(i).get("values")[4])))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +miktary.get()+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+str(table.item(i).get("values")[4])+"\n")                  
                                logsclose()


                            else:
                                table2.insert(parent = '', index='end',id=max([int(q) for q in table2.get_children()])+1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,miktary.get(),table.item(i).get("values")[3],vereny.get()
                                ,alany.get(),str(tarihy.get()),str(aciklamay.get())))

                                logs()
                                logfile.write("DEPO LİSTESİ"+"\t"+"ZİMMETLEME"+"     "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                                +miktary.get()+"  -  "+str(table.item(i).get("values")[3])+"  -  "+vereny.get()+"  -  "
                                +alany.get()+"  -  "+str(tarihy.get())+"  -  "+aciklamay.get()+"\n")                  
                                logsclose()
                    
                if(int(miktary.get())) >= int(table.item(i).get("values")[2]):
                    real_table3 = []
                    indexlist = []
                    table.delete(i)

                    for j in table.get_children():
                        real_table3.append(table.item(j).get("values"))
                        indexlist.append(j)
                    

                    for x in table.get_children():
                        table.delete(x)
                    
                    for l in range(0,len(real_table3)):    
                        table.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])
                        

                else:
                    table.item(i, text="",values=(table.item(i).get("values")[0],table.item(i).get("values")[1],table.item(i).get("values")[2]-int(miktary.get())
                    ,table.item(i).get("values")[3],table.item(i).get("values")[4])) 
                flag = True
                break
    sorting(table)
    sorting(table2)

    export("local")
    export2("local")

    if flag == False:
        messagebox.showwarning("UYARI","Ürün Bulunamadı!")

############################################################

############################################################
# Selecting product with mouse on registered table part.
def selectRegistered():
    try:

        miktary2.delete(0,END)
        saciklama.delete(0,END)
        global selected1
        selected1 = table.focus()
        smalzeme.place(x=100,y=225)
        smalzeme.configure(text=table.item(selected1).get("values")[0],fg="red")

        smalzeme_metin.place(x=100,y=250)
        smalzeme_metin.configure(text=table.item(selected1).get("values")[1],fg="red")

        miktary2.insert(0,table.item(selected1).get("values")[2])

        solcu.place(x=100,y=300)
        solcu.configure(text=table.item(selected1).get("values")[3],fg="red")

        saciklama.insert(0,table.item(selected1).get("values")[4])

    except IndexError:
        messagebox.showwarning("UYARI","Lütfen Ürün Seçin")

############################################################

############################################################
# cleaning and back button
def clearagain():
    malzemey.delete(0,END)
    miktary.delete(0,END)
    vereny.delete(0,END)
    alany.delete(0,END)
    tarihy.delete(0,END)
    aciklamay.delete(0,END)

    miktary2.delete(0,END)
    vereny2.delete(0,END)
    alany2.delete(0,END)
    tarihy2.delete(0,END)

    smalzeme.configure(text="")
    smalzeme_metin.configure(text="")
    solcu.configure(text="")
    saciklama.delete(0,END)


def backagain():
    malzemey.delete(0,END)
    miktary.delete(0,END)
    vereny.delete(0,END)
    alany.delete(0,END)
    tarihy.delete(0,END)
    aciklamay.delete(0,END)
    miktary2.delete(0,END)
    vereny2.delete(0,END)
    alany2.delete(0,END)
    tarihy2.delete(0,END)

    smalzeme.configure(text="")
    smalzeme_metin.configure(text="")
    solcu.configure(text="")
    saciklama.delete(0,END)

    remove_frame.place_forget()
    registered_remove_frame.place_forget()

    add_button.place(x=120,y=75)
    remove_button.place(x=120,y=225)
    edit_button.place(x=120,y=300)
    registered_removeitem_button.place(x=120,y=150)
############################################################

############################################################
# Interface of adding registered product back to inventory from registered list.
def removeRegistered():

    add_button.place_forget()
    remove_button.place_forget()
    edit_button.place_forget()
    registered_removeitem_button.place_forget()
    
    remove_frame.place(x=0,y=0)
    registered_remove_frame.place(x=0,y=0)

    malzeme_label = Label(registered_remove_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=25)
    malzeme_label = Label(registered_remove_frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=50)
    malzeme_label = Label(registered_remove_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=75)
    malzeme_label = Label(registered_remove_frame,text="Teslim Eden: ",bg="LavenderBlush2").place(x=5,y=100)
    malzeme_label = Label(registered_remove_frame,text="Teslim Alan: ",bg="LavenderBlush2").place(x=5,y=125)
    malzeme_label = Label(registered_remove_frame,text="Teslim Tarihi: ",bg="LavenderBlush2").place(x=5,y=150)

    miktary2.place(x=95,y=275)
    malzemey.place(x=95,y=25)
    miktary.place(x=95,y=50)
    aciklamay.place(x=95,y=75)

    vereny.place(x=95,y=100)
    alany.place(x=95,y=125)
    tarihy.place(x=95,y=150)
    
    ekley.place(x=250,y=75)
    sec.place(x=250,y=272)

    malzeme_label = Label(registered_remove_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=225)
    malzeme_label = Label(registered_remove_frame,text="Malzeme Metni: ",bg="LavenderBlush2").place(x=5,y=250)
    malzeme_label = Label(registered_remove_frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=275)
    malzeme_label = Label(registered_remove_frame,text="Ölçü Birimi: ",bg="LavenderBlush2").place(x=5,y=300)
    malzeme_label = Label(registered_remove_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=325)

    malzeme_label = Label(registered_remove_frame,text="Teslim Eden: ",bg="LavenderBlush2").place(x=5,y=360)
    malzeme_label = Label(registered_remove_frame,text="Teslim Alan: ",bg="LavenderBlush2").place(x=5,y=385)
    malzeme_label = Label(registered_remove_frame,text="Teslim Tarihi: ",bg="LavenderBlush2").place(x=5,y=410)

    secmeli = Label(registered_remove_frame,text="SEÇEREK ZİMMETLE",bg="LavenderBlush2",fg="indian red")
    secmeli.place(x=110,y=200)
    secmeli.config(font=("Calibri",12))

    secmeli2 = Label(registered_remove_frame,text="NO İLE ZİMMETLE",bg="LavenderBlush2",fg="indian red")
    secmeli2.place(x=120,y=0)
    secmeli2.config(font=("Calibri",12))
    
    malzeme_label = Label(registered_remove_frame,text="_____________________________________________________________________________ ",bg="LavenderBlush2").place(x=2,y=180)

    vereny2.place(x=95,y=360)
    alany2.place(x=95,y=385)
    tarihy2.place(x=95,y=410)
    saciklama.place(x=95,y=325)

    ekley2.place(x=250,y=320)
    geriy.place(x=250,y=355)
    temizley.place(x=250,y=390)

    updateTable(table,"Depo Listesi.xlsx")    
############################################################    

############################################################
# Editing registered product selecting with mouse.
def sec_fr():
    malzB.place_forget()
    secB.place_forget()
    geriB.place_forget()

    sec_Frame.place(x=0,y=0)

    sect.place(x=75,y=60)

    malzeme_label = Label(sec_Frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=125)
    malzeme_label = Label(sec_Frame,text="Malzeme Metni: ",bg="LavenderBlush2").place(x=5,y=150)
    malzeme_label = Label(sec_Frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=175)
    malzeme_label = Label(sec_Frame,text="Ölçü: ",bg="LavenderBlush2").place(x=5,y=200)
    malzeme_label = Label(sec_Frame,text="Teslim Eden: ",bg="LavenderBlush2").place(x=5,y=225)
    malzeme_label = Label(sec_Frame,text="Teslim Alan: ",bg="LavenderBlush2").place(x=5,y=250)
    malzeme_label = Label(sec_Frame,text="Tarih: ",bg="LavenderBlush2").place(x=5,y=275)
    malzeme_label = Label(sec_Frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=300)

    malzemet.place(x=100,y=125)
    malzeme_metint.place(x=100,y=150)
    miktart.place(x=100,y=175)
    olcut.place(x=100,y=200)
    verent.place(x=100,y=225)
    alant.place(x=100,y=250)
    tariht.place(x=100,y=275)
    aciklamat.place(x=100,y=300)

    duzenlet.place(x=35,y=350)
    gerit.place(x=130,y=350)
    temizlet.place(x=85,y=400)
    updateTable(table2,"Zimmet Listesi.xlsx")
############################################################

############################################################
# Cleaning and back buttons.
def clearEditto():
    malzemeq.delete(0,END)
    malzemeq2.delete(0,END)
    malzeme_metinq.delete(0,END)
    miktarq.delete(0,END)
    olcuq.delete(0,END)
    verenq.delete(0,END)
    alanq.delete(0,END)
    tarihq.delete(0,END)
    aciklamaq.delete(0,END)

    duzenleq['state'] = DISABLED
    malzemeq['state'] = NORMAL

def clearEditto2():
    malzemet.delete(0,END)
    malzeme_metint.delete(0,END)
    miktart.delete(0,END)
    olcut.delete(0,END)
    verent.delete(0,END)
    alant.delete(0,END)
    tariht.delete(0,END)
    aciklamat.delete(0,END)

    duzenlet['state'] = DISABLED

def backEditto():
    malzemeq['state'] = NORMAL
    duzenleq['state'] = DISABLED
    malzemeq.delete(0,END)
    malzemeq2.delete(0,END)
    malzeme_metinq.delete(0,END)
    miktarq.delete(0,END)
    olcuq.delete(0,END)
    verenq.delete(0,END)
    alanq.delete(0,END)
    tarihq.delete(0,END)
    aciklamaq.delete(0,END)

    malzNo_frame.place_forget()
    edit_register_frame.place(x=0,y=0)
    
    malzB.place(x=50,y=150)
    secB.place(x=50,y=225)
    geriB.place(x=61,y=375)
def backEditto2():

    malzemet.delete(0,END)
    malzeme_metint.delete(0,END)
    miktart.delete(0,END)
    olcut.delete(0,END)
    verent.delete(0,END)
    alant.delete(0,END)
    tariht.delete(0,END)
    aciklamat.delete(0,END)
    duzenlet['state'] = DISABLED
    sec_Frame.place_forget()
    edit_register_frame.place(x=0,y=0)

    malzB.place(x=50,y=150)
    secB.place(x=50,y=225)
    geriB.place(x=61,y=375)
############################################################

############################################################
# editing registered product with searching its product id.
def malzNo():
    malzB.place_forget()
    secB.place_forget()
    geriB.place_forget()

    malzNo_frame.place(x=0,y=0)

    malzeme_label = Label(malzNo_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=25)
    malzemeq.place(x=100,y=25)
    bulq.place(x=75,y=60)

    malzeme_label = Label(malzNo_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=125)
    malzeme_label = Label(malzNo_frame,text="Malzeme Metni: ",bg="LavenderBlush2").place(x=5,y=150)
    malzeme_label = Label(malzNo_frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=175)
    malzeme_label = Label(malzNo_frame,text="Ölçü: ",bg="LavenderBlush2").place(x=5,y=200)
    malzeme_label = Label(malzNo_frame,text="Teslim Eden: ",bg="LavenderBlush2").place(x=5,y=225)
    malzeme_label = Label(malzNo_frame,text="Teslim Alan: ",bg="LavenderBlush2").place(x=5,y=250)
    malzeme_label = Label(malzNo_frame,text="Tarih: ",bg="LavenderBlush2").place(x=5,y=275)
    malzeme_label = Label(malzNo_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=300)
    malzeme_label = Label(malzNo_frame,text="______________________________________________ ",bg="LavenderBlush2").place(x=0,y=90)

    malzemeq2.place(x=100,y=125)
    malzeme_metinq.place(x=100,y=150)
    miktarq.place(x=100,y=175)
    olcuq.place(x=100,y=200)
    verenq.place(x=100,y=225)
    alanq.place(x=100,y=250)
    tarihq.place(x=100,y=275)
    aciklamaq.place(x=100,y=300)

    duzenleq.place(x=35,y=350)
    geriq.place(x=130,y=350)
    temizleq.place(x=85,y=400)

    updateTable(table2,"Zimmet Listesi.xlsx")

############################################################

############################################################
# back button
def backRegis():
    edit_register_frame.place_forget()
############################################################

############################################################
# interface of edit screen.
def editRegister():

    edit_register_frame.place(x=0,y=0)
    malzB.place(x=50,y=150)
    secB.place(x=50,y=225)
    geriB.place(x=61,y=375)

    updateTable(table2,"Zimmet Listesi.xlsx")
############################################################

############################################################
# interface of deleting registered product
def removeRegister():
    remove_register_frame.place(x=0,y=0)
    malzeme_label = Label(remove_register_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=5,y=25)
    malzeme_label = Label(remove_register_frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=60)
    malzeme_label = Label(remove_register_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=95)

    malzemer.place(x=100,y=27)
    miktarr.place(x=100,y=62)
    hurdaaciklama3.place(x=100,y=97)
    cikarr.place(x=80,y=135)

    malzeme_label = Label(remove_register_frame,text="Miktar: ",bg="LavenderBlush2").place(x=5,y=220)
    malzeme_label = Label(remove_register_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=5,y=255)

    secmeli = Label(remove_register_frame,text="SEÇEREK SİL",bg="LavenderBlush2",fg="indian red")
    secmeli.place(x=80,y=182)
    secmeli.config(font=("Calibri",12))


    secmeli2 = Label(remove_register_frame,text="NO İLE SİL",bg="LavenderBlush2",fg="indian red")
    secmeli2.place(x=85,y=0)
    secmeli2.config(font=("Calibri",12))

    malzeme_label2 = Label(remove_register_frame,text="______________________________________________ ",bg="LavenderBlush2")
    malzeme_label2.place(x=2,y=161)
    
    miktarr2.place(x=100,y=222)
    hurdaaciklama4.place(x=100,y=257)

    cikarr2.place(x=15,y=305)
    gerir.place(x=145,y=305)
    temizler.place(x=85,y=355)
    
    updateTable(table2,"Zimmet Listesi.xlsx")

############################################################

############################################################
# editing nonregistered product on inventroy by selecting with mouse.
def edit2():

    logs()
    logfile.write("DEPO LİSTESİ"+"\t"+"DÜZENLEME"+"      "+str(table.item(selectedItem).get("values")[0])+"  -  "+str(table.item(selectedItem).get("values")[1])+"  -  "
    +str(table.item(selectedItem).get("values")[2])+"  -  "+str(table.item(selectedItem).get("values")[3])+"  -  "+str(table.item(selectedItem).get("values")[4]))

    table.item(selectedItem, text="",values=(malzeme_no3.get(),malzeme_metin3.get()
            ,miktar3.get(),olcu3.get(),aciklama3.get()))

    logfile.write("   -------YENİ ÜRÜN:-------   "+malzeme_no3.get()+"  -  "+malzeme_metin3.get()+"  -  "
    +miktar3.get()+"  -  "+olcu3.get()+"  -  "+aciklama3.get()+"\n")                  
    logsclose()


    sorting(table)

    export("local")
############################################################
    
############################################################
# cleaning and back buttons.
def clearEdit2():
    malzeme_no3.delete(0,END)
    malzeme_metin3.delete(0,END)
    miktar3.delete(0,END)
    olcu3.delete(0,END)
    aciklama3.delete(0,END)
    editB2['state'] = DISABLED

def backEdit():

    malzeme_no3.delete(0,END)
    malzeme_metin3.delete(0,END)
    miktar3.delete(0,END)
    olcu3.delete(0,END)
    aciklama3.delete(0,END)
    
    registered_edit_frame.place_forget()
    editB2['state'] = DISABLED
    notregistered_edit_button.place(x=145,y=125)
    registered_edit_button.place(x=146,y=225)
############################################################

############################################################
# selecting product with mouse to edit it.
def findSelected():
    try:
        global selectedItem
        selectedItem = table.focus()
        

        malzeme_no3.delete(0,END)
        malzeme_metin3.delete(0,END)
        miktar3.delete(0,END)
        olcu3.delete(0,END)
        aciklama3.delete(0,END)

        malzeme_no3.insert(0,str(table.item(selectedItem).get("values")[0]))
        malzeme_metin3.insert(0,str(table.item(selectedItem).get("values")[1]))
        miktar3.insert(0,str(table.item(selectedItem).get("values")[2]))
        olcu3.insert(0,str(table.item(selectedItem).get("values")[3]))
        aciklama3.insert(0,str(table.item(selectedItem).get("values")[4]))

        editB2['state'] = ACTIVE
    except IndexError:
        messagebox.showwarning("UYARI","Lütfen Ürün Seçin!")
############################################################

############################################################
# interface of edit screen on inventory.
def selectedEdit():
    notregistered_edit_button.place_forget()
    registered_edit_button.place_forget()
    registered_edit_frame.place(x=0,y=0)

    no_malzeme3 = Label(registered_edit_frame,text="Malzeme No: ",bg="LavenderBlush2").place(x=25,y=100)
    malzeme_no3.place(x=125,y=102)

    malzeme_met = Label(registered_edit_frame,text="Malzeme Metni: ",bg="LavenderBlush2").place(x=25,y=140)
    malzeme_metin3.place(x=125,y=142)

    miktar3_ = Label(registered_edit_frame,text="Miktar: ",bg="LavenderBlush2").place(x=25,y=180)
    miktar3.place(x=125,y=182)

    olcu3_ = Label(registered_edit_frame,text="Ölçü Birimi: ",bg="LavenderBlush2").place(x=25,y=220)
    olcu3.place(x=125,y=222)

    aciklama3_ = Label(registered_edit_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=25,y=260)
    aciklama3.place(x=125,y=262,width=225)

    searchB2.place(x=138,y=40)
    editB2.place(x=75,y=300)
    backB2.place(x=200,y=300)
    clearB2.place(x=138,y=350)

    updateTable(table,"Depo Listesi.xlsx")
############################################################

############################################################
# editing nonregistered product by searching it product id.
def edit1():

    logs()
    logfile.write("DEPO LİSTESİ"+"\t"+"DÜZENLEME"+"      "+str(table.item(productE).get("values")[0])+"  -  "+str(table.item(productE).get("values")[1])+"  -  "
    +str(table.item(productE).get("values")[2])+"  -  "+str(table.item(productE).get("values")[3])+"  -  "+str(table.item(productE).get("values")[4]))                  
    
    table.item(productE, text="",values=(malzeme_no2.get(),malzeme_metin2.get()
            ,miktar2.get(),olcu2.get(),aciklama2.get()))
    
    logfile.write("   -------YENİ ÜRÜN:-------   "+malzeme_no2.get()+"  -  "+malzeme_metin2.get()+"  -  "
    +miktar2.get()+"  -  "+olcu2.get()+"  -  "+aciklama2.get()+"\n")                  
    logsclose()

    
    sorting(table)
    
    export("local")
############################################################

############################################################
# cleaning and back buttons.
def clearEdit():
    malzeme_no1.delete(0,END)
    malzeme_no2.delete(0,END)
    malzeme_metin2.delete(0,END)
    miktar2.delete(0,END)
    olcu2.delete(0,END)
    aciklama2.delete(0,END)
    editB['state'] = DISABLED
    malzeme_no1['state'] = NORMAL

def backToEdit():
    
    malzeme_no1['state'] = NORMAL
    malzeme_no1.delete(0,END)
    malzeme_no2.delete(0,END)
    malzeme_metin2.delete(0,END)
    miktar2.delete(0,END)
    olcu2.delete(0,END)
    aciklama2.delete(0,END)
    editB['state'] = DISABLED
    notregistered_edit_frame.place_forget()

    notregistered_edit_button.place(x=142,y=145)
    registered_edit_button.place(x=143,y=245)
############################################################

############################################################
# searching entered product id by user on the table to edit.
def search():

    ctrl = False
    global productE
    productE = 0
    for i in table.get_children():
        if(malzeme_no1.get() == str(table.item(i).get("values")[0])):
            productE = i
            malzeme_no1['state'] = DISABLED

            malzeme_no2.delete(0,END)
            malzeme_metin2.delete(0,END)
            miktar2.delete(0,END)
            olcu2.delete(0,END)
            aciklama2.delete(0,END)

            editB['state'] = ACTIVE

            malzeme_no2.insert(0,str(table.item(i).get("values")[0]))
            malzeme_metin2.insert(0,str(table.item(i).get("values")[1]))
            miktar2.insert(0,str(table.item(i).get("values")[2]))
            olcu2.insert(0,str(table.item(i).get("values")[3]))
            aciklama2.insert(0,str(table.item(i).get("values")[4]))
            ctrl = True
            break
        
    if ctrl == False:
        messagebox.showwarning("UYARI","Ürün Mevcut Değil!")
############################################################            

############################################################
# interface of edit screen on inventory.
def no():
    notregistered_edit_button.place_forget()
    registered_edit_button.place_forget()
    notregistered_edit_frame.place(x=0,y=0)
    

    no_malzeme = Label(notregistered_edit_frame,text="Malzeme No: ",bg="LavenderBlush2").place(x=25,y=20)
    malzeme_no1.place(x=125,y=22)

    no_malzeme2 = Label(notregistered_edit_frame,text="Malzeme No: ",bg="LavenderBlush2").place(x=25,y=100)
    malzeme_no2.place(x=125,y=102)

    malzeme_met = Label(notregistered_edit_frame,text="Malzeme Metni: ",bg="LavenderBlush2").place(x=25,y=140)
    malzeme_metin2.place(x=125,y=142)

    miktar2_ = Label(notregistered_edit_frame,text="Miktar: ",bg="LavenderBlush2").place(x=25,y=180)
    miktar2.place(x=125,y=182)

    olcu2_ = Label(notregistered_edit_frame,text="Ölçü Birimi: ",bg="LavenderBlush2").place(x=25,y=220)
    olcu2.place(x=125,y=222)

    aciklama2_ = Label(notregistered_edit_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=25,y=260)
    aciklama2.place(x=125,y=262,width=225)

    malzeme_label = Label(notregistered_edit_frame,text="_____________________________________________________________________________ ",bg="LavenderBlush2").place(x=2,y=55)

    searchB.place(x=265,y=18)
    editB.place(x=75,y=300)
    backB.place(x=200,y=300)
    clearB.place(x=138,y=350)

    updateTable(table,"Depo Listesi.xlsx")
############################################################

############################################################
# cleaning button
def clear_remove():
    malzeme_no.delete(0,END)
    sayi.delete(0,END)
    sayi2.delete(0,END)
    hurdaaciklama.delete(0,END)
    hurdaaciklama2.delete(0,END)
############################################################    

############################################################
# deleting nonregistered item by selecting with mouse.
def decrease():
    
    try:
    
        selected = table.focus()
        
        if sayi2.get().isspace() or sayi2.get() == '':
            if table.item(selected).get("values")[2]<=1:
                try:
                    table3.insert(parent = '', index=0,id=1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                                ,1,table.item(selected).get("values")[3],str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))
                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                                ,1,table.item(selected).get("values")[3],str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))
                
                logs()
                logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(selected).get("values")[0])+"  -  "+str(table.item(selected).get("values")[1])+"  -  "
                +"1"+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+hurdaaciklama2.get()+"\n")                  
                logsclose()


                real_table3 = []
                indexlist = []
                table.delete(selected)

                for i in table.get_children():
                    real_table3.append(table.item(i).get("values"))
                    indexlist.append(i)


                for x in table.get_children():
                    table.delete(x)

                
                for l in range(0,len(real_table3)):    
                    table.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])
                
            else:

                try:
                    table3.insert(parent = '', index=0,id=1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                            ,1,table.item(selected).get("values")[3],str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))
                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                            ,1,table.item(selected).get("values")[3],str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))

                logs()
                logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(selected).get("values")[0])+"  -  "+str(table.item(selected).get("values")[1])+"  -  "
                +"1"+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+hurdaaciklama2.get()+"\n")                  
                logsclose()
                
                table.item(selected, text="",values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1],table.item(selected).get("values")[2]-1
                ,table.item(selected).get("values")[3],table.item(selected).get("values")[4]))       
        else:
            if(table.item(selected).get("values")[2] - int(sayi2.get())) <=0:

                try:
                    table3.insert(parent = '', index=0,id=1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                            ,int(table.item(selected).get("values")[2]),table.item(selected).get("values")[3]
                            ,str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(selected).get("values")[0])+"  -  "+str(table.item(selected).get("values")[1])+"  -  "
                    +str(table.item(selected).get("values")[2])+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+hurdaaciklama2.get()+"\n")                  
                    logsclose()


                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                            ,int(table.item(selected).get("values")[2]),table.item(selected).get("values")[3]
                            ,str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(selected).get("values")[0])+"  -  "+str(table.item(selected).get("values")[1])+"  -  "
                    +str(table.item(selected).get("values")[2])+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+hurdaaciklama2.get()+"\n")                  
                    logsclose()
                
                real_table4 = []
                indexlist2 = []
                table.delete(selected)

                for i in table.get_children():
                    real_table4.append(table.item(i).get("values"))
                    indexlist2.append(i)


                for x in table.get_children():
                    table.delete(x)

                
                for l in range(0,len(real_table4)):    
                    table.insert(parent = '', index='end',id=indexlist2[l],values=[real_table4[l][a] for a in range (0,len(real_table4[0]))])
                
            else:

                try:
                    table3.insert(parent = '', index=0,id=1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                            ,int(sayi2.get()),table.item(selected).get("values")[3],str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(selected).get("values")[0])+"  -  "+str(table.item(selected).get("values")[1])+"  -  "
                    +sayi2.get()+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+hurdaaciklama2.get()+"\n")                  
                    logsclose()


                except:
                    table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                            ,int(sayi2.get()),table.item(selected).get("values")[3],str(str(hurdaaciklama2.get())+"     "+str(datetime.now())[0:19])))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(selected).get("values")[0])+"  -  "+str(table.item(selected).get("values")[1])+"  -  "
                    +sayi2.get()+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+hurdaaciklama2.get()+"\n")                  
                    logsclose()

                table.item(selected, text="",values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
                ,table.item(selected).get("values")[2]-int(sayi2.get())
                ,table.item(selected).get("values")[3],table.item(selected).get("values")[4]))
        
        sorting(table)

        export("local")
        export3("local")
    except IndexError:
        messagebox.showwarning("UYARI","Lütfen Ürün Seçin!") 
############################################################

############################################################
# deleting nonregistered item by searching it with product id.
def removeNotRegisteredItem():
    global real_table2
    real_table2 = []

    flag = False
    
    for j in table.get_children():

        if(malzeme_no.get() == str(table.item(j).get("values")[0])):
            flag = True
            break

    if flag == True:
        for i in table.get_children():
            
                if(malzeme_no.get().isspace() or malzeme_no.get() == '' or sayi.get().isspace() or sayi.get() == ''):
                    
                    messagebox.showwarning("UYARI","Parametreyi lütfen doldurun!")
                    break
                
                elif(str(table.item(i).get("values")[0]) == malzeme_no.get()):
                 
                    if(table.item(i).get("values")[2] - int(sayi.get())) <=0:
                        try:
                            table3.insert(parent = '', index=0,id=1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,int(table.item(i).get("values")[2]),table.item(i).get("values")[3],str(str(hurdaaciklama.get())+"     "+str(datetime.now())[0:19])))
                        except:
                            table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,int(table.item(i).get("values")[2]),table.item(i).get("values")[3],str(str(hurdaaciklama.get())+"     "+str(datetime.now())[0:19])))
                        

                        logs()
                        logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                        +str(table.item(i).get("values")[2])+"  -  "+str(table.item(i).get("values")[3])+"  -  "+hurdaaciklama.get()+"\n")                  
                        logsclose()



                        real_table3 = []
                        indexlist = []
                        table.delete(i)

                        for j in table.get_children():
                            real_table3.append(table.item(j).get("values"))
                            indexlist.append(j)
                        

                        for x in table.get_children():
                            table.delete(x)
                        
                        for l in range(0,len(real_table3)):    
                            table.insert(parent = '', index='end',id=indexlist[l],values=[real_table3[l][a] for a in range (0,len(real_table3[0]))])
                
                    
                    else:
                        try:
                            table3.insert(parent = '', index=0,id=1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,int(sayi.get()),table.item(i).get("values")[3],str(str(hurdaaciklama.get())+"     "+str(datetime.now())[0:19])))
                        except:
                            table3.insert(parent = '', index=0,id=max([int(q) for q in table3.get_children()])+1,values=(table.item(i).get("values")[0],table.item(i).get("values")[1]
                                ,int(sayi.get()),table.item(i).get("values")[3],str(str(hurdaaciklama.get())+"     "+str(datetime.now())[0:19])))

                        logs()
                        logfile.write("DEPO LİSTESİ"+"\t"+"SİLME"+"          "+str(table.item(i).get("values")[0])+"  -  "+str(table.item(i).get("values")[1])+"  -  "
                        +sayi.get()+"  -  "+str(table.item(i).get("values")[3])+"  -  "+hurdaaciklama.get()+"\n")                  
                        logsclose()


                        table.item(i,values=(table.item(i).get("values")[0],table.item(i).get("values")[1],table.item(i).get("values")[2]
                        -int(sayi.get()),table.item(i).get("values")[3],table.item(i).get("values")[4]))
                        #table.item(i).get("values")[2] +=1
                    break

        sorting(table)

        export("local")
        export3("local")
    elif flag == False:
        messagebox.showwarning("UYARI","Böyle Bir Ürün Yok!")
############################################################

############################################################
# interface of deleting screen.
def removeNotRegistered():

    add_button.place_forget()
    remove_button.place_forget()
    edit_button.place_forget()
    registered_removeitem_button.place_forget()

    remove_frame.place(x=0,y=0)
    notregistered_remove_frame.place(x=0,y=0)
    

    malzeme_label = Label(notregistered_remove_frame,text="Malzeme NO: ",bg="LavenderBlush2").place(x=25,y=40)
    malzeme_no.place(x=125,y=40)

    miktar_label = Label(notregistered_remove_frame,text="Miktar: ",bg="LavenderBlush2").place(x=25,y=80)
    sayi.place(x=125,y=80)

    miktar_label = Label(notregistered_remove_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=25,y=120)
    hurdaaciklama.place(x=125,y=120)

    remove_notregistered.place(x=265,y=75)

    miktar2_label = Label(notregistered_remove_frame,text="Miktar: ",bg="LavenderBlush2").place(x=25,y=215)
    sayi2.place(x=125,y=217)

    miktar2_label = Label(notregistered_remove_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=25,y=255)
    hurdaaciklama2.place(x=125,y=257)

    secmeli = Label(notregistered_remove_frame,text="SEÇEREK SİL",bg="LavenderBlush2",fg="indian red")
    secmeli.place(x=150,y=180)
    secmeli.config(font=("Calibri",12))

    secmeli2 = Label(notregistered_remove_frame,text="NO İLE SİL",bg="LavenderBlush2",fg="indian red")
    secmeli2.place(x=150,y=0)
    secmeli2.config(font=("Calibri",12))
    
    malzeme_label = Label(notregistered_remove_frame,text="_____________________________________________________________________________ ",bg="LavenderBlush2").place(x=2,y=160)

    clear_remove2.place(x=150,y=360)

    remove_notregistered2.place(x=60,y=310)
    geri.place(x=215,y=310)

    updateTable(table,"Depo Listesi.xlsx")    
############################################################

############################################################
# back button
def backRemoveRegister():
    malzeme_no.delete(0, END)
    sayi.delete(0, END)
    sayi2.delete(0, END)
    hurdaaciklama.delete(0,END)
    hurdaaciklama2.delete(0,END)
    
    remove_frame.place_forget()
    notregistered_remove_frame.place_forget()

    add_button.place(x=120,y=75)
    remove_button.place(x=120,y=225)
    edit_button.place(x=120,y=300)
    registered_removeitem_button.place(x=120,y=150)
############################################################    

############################################################
# interface of adding nonregistered item to table.
def addNotRegistered():

    add_button.place_forget()
    remove_button.place_forget()
    edit_button.place_forget()

    add_frame.place(x=0,y=0),

    notregistered_item_frame.place(x=0,y=0)
    registered_removeitem_button.place_forget()

    malzeme_label = Label(notregistered_item_frame,text="Malzeme No: ",bg="LavenderBlush2").place(x=50,y=25)
    malzeme.place(x=50,y=50)

    malzeme_metin_label = Label(notregistered_item_frame,text="Malzeme metni: ",bg="LavenderBlush2").place(x=50,y=80)
    malzeme_metin.place(x=50,y=105)

    miktar_label = Label(notregistered_item_frame,text="Adet: ",bg="LavenderBlush2").place(x=50,y=135)
    miktar.place(x=50,y=160)

    olcu_label = Label(notregistered_item_frame,text="Ölçü: ",bg="LavenderBlush2").place(x=50,y=190)
    olcu.place(x=50,y=215)

    aciklama_label = Label(notregistered_item_frame,text="Açıklama: ",bg="LavenderBlush2").place(x=50,y=245)
    aciklama.place(x=50,y=270)

    aciklama_label = Label(notregistered_item_frame,text="_____________________________________________________________________________ ",bg="LavenderBlush2").place(x=2,y=380)

    artma_label = Label(notregistered_item_frame,text="Adet: ",bg="LavenderBlush2").place(x=202,y=407)
    artma.place(x=245,y=408)

    add_notregistered_button.place(x=75,y=300)
    back_button.place(x=200,y=300)
    clear_button.place(x=137,y=350)
    increase_button.place(x=65,y=405)
    
    updateTable(table,"Depo Listesi.xlsx")
############################################################

############################################################
# back button
def backToRegister():
    malzeme.delete(0, END)
    malzeme_metin.delete(0, END)
    miktar.delete(0,END)
    olcu.delete(0, END)
    aciklama.delete(0, END)
    artma.delete(0, END)


    notregistered_item_frame.place_forget()
    add_frame.place_forget()

    add_button.place(x=120,y=75)
    remove_button.place(x=120,y=225)
    edit_button.place(x=120,y=300)
    registered_removeitem_button.place(x=120,y=150)
############################################################

############################################################
# adding nonregistered item to inventory
def addNotRegisteredItem():

    
    flag = False

    try:
        for i in table.get_children():
            if(malzeme.get().isspace() or malzeme.get() == '' or malzeme_metin.get().isspace() or malzeme_metin.get() == '' 
            or olcu.get().isspace() or olcu.get() == ''):
                flag = True
                
                messagebox.showwarning("UYARI","Lütfen parametreleri doldurun!")
                break

            elif(str(table.item(i).get("values")[0]) == malzeme.get() and str(table.item(i).get("values")[1]) == malzeme_metin.get() and
            str(table.item(i).get("values")[3]) == olcu.get() and str(table.item(i).get("values")[4]) == aciklama.get()):

                flag = True
                
                
                if miktar.get().isspace() or miktar.get() == '':
                    table.item(i,values=(table.item(i).get("values")[0],table.item(i).get("values")[1],table.item(i).get("values")[2]+1
                    ,table.item(i).get("values")[3],table.item(i).get("values")[4]))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+malzeme.get()+"  -  "+malzeme_metin.get()+"  -  "
                    +"1"+"  -  "+olcu.get()+"  -  "+aciklama.get()+"\n")                   
                    logsclose()

                    break
                else:
                    table.item(i,values=(table.item(i).get("values")[0],table.item(i).get("values")[1],
                    table.item(i).get("values")[2]+int(miktar.get())
                    ,table.item(i).get("values")[3],table.item(i).get("values")[4]))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+malzeme.get()+"  -  "+malzeme_metin.get()+"  -  "
                    +miktar.get()+"  -  "+olcu.get()+"  -  "+aciklama.get()+"\n")                   
                    logsclose()
                    break

    except ValueError:
        messagebox.showwarning("UYARI","Lütfen Miktarı sayı olarak giriniz!")               

    if(flag == False):
        
        
            if miktar.get().isspace() or miktar.get() == '':
                try:
                    table.insert(parent='',index='end',id = max([int(q) for q in table.get_children()])+1,values = (malzeme.get(),malzeme_metin.get(),1,olcu.get(),aciklama.get()))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+malzeme.get()+"  -  "+malzeme_metin.get()+"  -  "
                    +"1"+"  -  "+olcu.get()+"  -  "+aciklama.get()+"\n")
                    logsclose()

                except:
                    table.insert(parent='',index='end',id = 1,values = (malzeme.get(),malzeme_metin.get(),1,olcu.get(),aciklama.get()))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+malzeme.get()+"  -  "+malzeme_metin.get()+"  -  "+"1"+"  -  "
                    +olcu.get()+"  -  "+aciklama.get()+"\n")
                    logsclose()

            else:
                try:
                    table.insert(parent='',index='end',id = max([int(q) for q in table.get_children()])+1,values = (malzeme.get(),malzeme_metin.get(),int(miktar.get()),olcu.get(),aciklama.get()))

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+malzeme.get()+"  -  "+malzeme_metin.get()
                    +"  -  "+miktar.get()+"  -  "+olcu.get()+"  -  "+aciklama.get()+"\n")                   
                    logsclose()
                except:
                    table.insert(parent='',index='end',id = 1,values = (malzeme.get(),malzeme_metin.get(),int(miktar.get()),olcu.get(),aciklama.get()))  

                    logs()
                    logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+malzeme.get()+"  -  "+malzeme_metin.get()+"  -  "+miktar.get()
                    +"  -  "+olcu.get()+"  -  "+aciklama.get()+"  -  ")                   
                    logsclose()           
    
    sorting(table)
    export("local")
############################################################    

############################################################
# increasing amount of selected item on the inventory.
def increase():
    
    try:
        selected = table.focus()
        
        if artma.get().isspace() or artma.get() == '':
            table.item(selected, text="",values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1],table.item(selected).get("values")[2]+1
                ,table.item(selected).get("values")[3],table.item(selected).get("values")[4]))

            logs()
            logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+str(table.item(selected).get("values")[0])+"  -  "+str((table.item(selected).get("values")[1]))+"  -  "
            +"1"+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+str(table.item(selected).get("values")[4])+"\n")                   
            logsclose()

        else:
            table.item(selected, text="",values=(table.item(selected).get("values")[0],table.item(selected).get("values")[1]
            ,table.item(selected).get("values")[2]+int(artma.get())
            ,table.item(selected).get("values")[3],table.item(selected).get("values")[4]))

            logs()
            logfile.write("DEPO LİSTESİ"+"\t"+"EKLEME"+"         "+str(table.item(selected).get("values")[0])+"  -  "+str((table.item(selected).get("values")[1]))+"  -  "
            +artma.get()+"  -  "+str(table.item(selected).get("values")[3])+"  -  "+str(table.item(selected).get("values")[4])+"\n")                   
            logsclose()

        export("local")

    except IndexError:
        messagebox.showwarning("UYARI","Lütfen Ürün Seçin!")  
############################################################
# clear button      
def clear():
    malzeme.delete(0, END)
    malzeme_metin.delete(0, END)
    miktar.delete(0,END)
    olcu.delete(0, END)
    aciklama.delete(0, END)
    artma.delete(0, END)
############################################################

############################################################
# interface of removing nonregistered item 
def removeItemFrame(inventory_frame,inventory_frame2,remove_frame,add_button,remove_button,edit_button):
    add_button.place_forget()
    remove_button.place_forget()
    registered_removeitem_button.place_forget()
    edit_button.place_forget()
    notregistered_remove_frame.place(x=0,y=0)
############################################################

############################################################
# interface of editing nonregistered item.   
def editFrame(inventory_frame,inventory_frame2,edit_frame,add_button,remove_button,edit_button,editBack_item_button):
    add_button.place_forget()
    remove_button.place_forget()
    edit_button.place_forget()
    registered_removeitem_button.place_forget()
    
    inventory_frame.place(x=0,y=0)
    inventory_frame2.place(x=950,y=100)
    
    edit_frame.place(x=0,y=0)
    
    
    editBack_item_button.place(x=160,y=400)

    notregistered_edit_button.place(x=142,y=145)
    registered_edit_button.place(x=143,y=245)
    
    updateTable(table,"Depo Listesi.xlsx")
############################################################

############################################################
# back button
def back(add_button,remove_button,edit_button,add_frame,remove_frame,edit_frame):
    add_frame.place_forget()
    remove_frame.place_forget()
    edit_frame.place_forget()
    
    add_button.place(x=120,y=75)
    remove_button.place(x=120,y=225)
    edit_button.place(x=120,y=300)
    registered_removeitem_button.place(x=120,y=150)

#########################################################################################################


#########################################################################################################        
# select inventory table to see it.
def showInventory(inventory_frame,inventory_frame2,register_frame,register_frame2):
    ## gizlenen
    inventory_frame.place_forget()
    inventory_frame2.place_forget()
    register_frame.place_forget()
    register_frame2.place_forget()
    junk_frame.place_forget()
    photo2.place_forget()

    
    ## görüntülenen
    inventory_frame.place(x=0,y=0)
    inventory_frame2.place(x=950,y=100)

    register_select.configure(bg="white")
    inventory_select.configure(bg="cadet blue")
    junk_select.configure(bg="white")
    
    global p, p2
    if p == 0 and p2 == 1:
        p2 = 0
        sorting(table)
        export("local")
        p2 = 1
    elif p == 0 and p2 == 0:
        sorting(table)
        export("local")

    updateTable(table,"Depo Listesi.xlsx")    
############################################################

############################################################
# select registered product table to see it.
def showRegister(inventory_frame,inventory_frame2,register_frame,register_frame2):
    ##gizlenen
    inventory_frame.place_forget()
    inventory_frame2.place_forget()
    register_frame.place_forget()
    register_frame2.place_forget()
    junk_frame.place_forget()
    photo2.place_forget()

    
    ## görüntülenen
    register_frame.place(x=0,y=0)
    register_frame2.place(x=1075,y=100)
    remove_product.place(x=64,y=125)
    edit_product.place(x=64,y=275)
    back_product.place(x=54,y=200)

    register_select.configure(bg="cadet blue")
    inventory_select.configure(bg="white")
    junk_select.configure(bg="white")

    global p,p2
    if p == 1 and p2 == 0:
        p = 0
        sorting(table2)
        export2("local")
        p = 1

    updateTable(table2,"Zimmet Listesi.xlsx")    
############################################################    

############################################################
# select junk table to see it.
def showJunk(inventory_frame,inventory_frame2,register_frame,register_frame2):
    inventory_frame.place_forget()
    inventory_frame2.place_forget()
    register_frame.place_forget()
    register_frame2.place_forget()
    
    junk_frame.place(x=0,y=0)
    photo2.place(x=1060,y=250)

    register_select.configure(bg="white")
    inventory_select.configure(bg="white")
    junk_select.configure(bg="cadet blue")
    
    updateTable(table3,"Hurda Listesi.xlsx")    
############################################################

############################################################
# saving inventory data to excel file.
def export(control):
    coloring()
    global p,p2

    if p != 1:
        global excel
        
        excel_row = list()
        for i in table.get_children():
                excel_row.append(table.item(i).get("values"))

        excel_column = list()

        for i in range(0,len(items.columns)):
            excel_column.append(table.heading(i)["text"])

        

        excel = pd.DataFrame(excel_row,columns=excel_column)

        if control == "server":
            print()
        elif control == "local":
            writer = pd.ExcelWriter("Tablolar/Depo Listesi.xlsx")

        
        excel.to_excel(writer,sheet_name='Envanter',index=False,na_rep ='NaN')

        for column in excel:
            column_width = max(excel[column].astype(str).map(len).max(), len(column))
            col_idx = excel.columns.get_loc(column)
            writer.sheets['Envanter'].set_column(col_idx, col_idx, column_width)
        
        col_idx = excel.columns.get_loc('Malzeme')
        writer.sheets['Envanter'].set_column(col_idx, col_idx, 15)

        col_idx = excel.columns.get_loc('Açıklama')
        writer.sheets['Envanter'].set_column(col_idx, col_idx, 50)

        writer.save()
    coloring()
############################################################

############################################################
# saving registered product table data to excel file.    
def export2(control):
    coloring2()
    global p,p2

    if p2 !=1:
        global excel2
        
        excel_row2 = list()
        for i in table2.get_children():
                excel_row2.append(table2.item(i).get("values"))

        excel_column2 = list()

        for i in range(0,len(items2.columns)):
            excel_column2.append(table2.heading(i)["text"])

        

        excel2 = pd.DataFrame(excel_row2,columns=excel_column2)
        
        excel2['Açıklama'] = excel2['Açıklama'].astype(str)
        if control == "server":
            print()
        elif control == "local":
            writer = pd.ExcelWriter("Tablolar/Zimmet Listesi.xlsx")
        
        excel2.to_excel(writer,sheet_name='Zimmet Listesi',index=False,na_rep ='NaN')

        for column in excel2:
            column_width = max(excel2[column].astype(str).map(len).max(), len(column))
            col_idx = excel2.columns.get_loc(column)
            writer.sheets['Zimmet Listesi'].set_column(col_idx, col_idx, column_width)
        
        col_idx = excel2.columns.get_loc('Malzeme')
        writer.sheets['Zimmet Listesi'].set_column(col_idx, col_idx, 15)

        col_idx = excel2.columns.get_loc('Açıklama')
        writer.sheets['Zimmet Listesi'].set_column(col_idx, col_idx, 50)

        writer.save()
    coloring2()
############################################################

############################################################
# saving junk table data to excel file.
def export3(control):
    coloring3()
    global p,p2

    if p3 !=1:
        global excel3
        
        excel_row3 = list()
        for i in table3.get_children():
                excel_row3.append(table3.item(i).get("values"))

        excel_column3 = list()
        

        for i in range(0,len(items.columns)):
            excel_column3.append(table3.heading(i)["text"])

        

        excel3 = pd.DataFrame(excel_row3,columns=excel_column3)
        if control == "server":
            print()
        elif control == "local":
            writer = pd.ExcelWriter("Tablolar/Hurda Listesi.xlsx")

        
        excel3.to_excel(writer,sheet_name='Hurda Listesi',index=False,na_rep ='NaN')

        for column in excel3:
            column_width = max(excel3[column].astype(str).map(len).max(), len(column))
            col_idx = excel3.columns.get_loc(column)
            writer.sheets['Hurda Listesi'].set_column(col_idx, col_idx, column_width)
        
        col_idx = excel3.columns.get_loc('Malzeme')
        writer.sheets['Hurda Listesi'].set_column(col_idx, col_idx, 15)

        col_idx = excel3.columns.get_loc('Açıklama')
        writer.sheets['Hurda Listesi'].set_column(col_idx, col_idx, 50)

        writer.save()
    coloring3()
############################################################

###########################################################               MAIN FUNCTION               #############################################################

def main():
    #main window
    global window    
    window = Tk()
    window.title("Envanter")
    window.geometry("1350x600")
    window.configure(bg="Lavender")
    window.resizable(False, False)
    
    # logos
    photo = Canvas(window,width=100, height=55, bg='Lavender',highlightthickness=0)
    logo = PhotoImage(file="Media/logo.png")
    photo.create_image(0, 0, image=logo, anchor=NW)
    photo.place(x=1255,y=555)
    
    window.iconbitmap('Media/simge.ico')
    
    global photo2
    photo2 = Canvas(window,width=275, height=130, bg='lavender',highlightthickness=0)
    logo2 = PhotoImage(file="Media/buyuklogo.png")
    photo2.create_image(0, 0, image=logo2, anchor=NW)
    photo2.place(x=1060,y=250)

       
    
    # Inventory screen
    inventory_frame = Frame(window,width=950,height=600,highlightbackground='black',highlightthickness=3)
    inventory_frame.place(x=0,y=0)
    
    # Registered product table screen
    inventory_frame2 = Frame(window,highlightbackground='black',width=400,height=450,bg="Lavender")
    inventory_frame2.place(x=950,y=100)
    
    global inventory_select
    inventory_select = Button(window,text = "Depo Listesi",width=10,bg="cadet blue",command= lambda : showInventory(inventory_frame,inventory_frame2,register_frame,register_frame2))
    inventory_select.place(x=1060,y=30)
    
    #frames of inventory screen
    global add_frame,remove_frame
    add_frame = Frame(inventory_frame2,width=400,height=445,bg="Lavender")
    remove_frame = Frame(inventory_frame2,width=400,height=445,bg="Lavender")
    edit_frame = Frame(inventory_frame2,width=400,height=445,bg="LavenderBlush2",highlightbackground='black',highlightthickness=1)
    
    global add_button,remove_button,edit_button
    
    #Buttons on inventory screen
    add_button = Button(inventory_frame2,text = 'Depoya Ürün Ekle',width=20,height = 2,bg="light steel blue",command = addNotRegistered)
    add_button.place(x=120,y=75)
    
    remove_button = Button(inventory_frame2,text = 'Depodan Ürün Sil',width=20,height =2,bg='light steel blue',command = removeNotRegistered)
    remove_button.place(x=120,y=225)
    
    edit_button = Button(inventory_frame2,text = 'Depodaki Ürünü Düzenle',width=20,height=2,bg='light steel blue',command = lambda: editFrame(inventory_frame,inventory_frame2,edit_frame,add_button,remove_button,edit_button,editBack_item_button))
    edit_button.place(x=120,y=300)

    global excel_button
    excel_button = Button(inventory_frame,text = 'Excele Aktar',width=20,height=2,bg = 'cadet blue',command= lambda: export("local"))
    excel_button.place(x=650,y=545)
    
    #Filter system of inventory screen
    global category_entry,filter_button,filter_cancel
    category_label = Label(inventory_frame,text = 'Ürün İsmi Giriniz: ')
    category_label.place(x=30,y=555)
    category_entry = Entry(inventory_frame,width = 25)
    category_entry.place(x=140,y=557)
    
    filter_button = Button(inventory_frame,text = 'Filtrele',width=20,height=2,bg='cadet blue',command=filter)
    filter_button.place(x=350,y=545)
    filter_cancel = Button(inventory_frame,text = 'X',width=2,height=2,bg='salmon',command=returnTable)
    filter_cancel.place(x=520,y=545)

    filter_cancel['state'] = DISABLED
    
    ## Second frame, adding product
    global registered_item_button
    global notregistered_item_frame
    notregistered_item_frame = Frame(add_frame,width=400,height=442,highlightbackground='black',highlightthickness=1,bg= "LavenderBlush2")
    

    

    ## Adding nonregistered product
    
    global malzeme, malzeme_metin, olcu, aciklama, miktar, artma

    malzeme = Entry(notregistered_item_frame)
    malzeme_metin = Entry(notregistered_item_frame)
    
    miktar = Entry(notregistered_item_frame,validate="key")
    miktar['validatecommand'] = (miktar.register(testVal),'%P','%d')
    
    olcu = Entry(notregistered_item_frame)
    aciklama = Entry(notregistered_item_frame)

    artma = Entry(notregistered_item_frame,width=10,validate="key")
    artma['validatecommand'] = (artma.register(testVal),'%P','%d')
    
    global add_notregistered_button, back_button, clear_button, increase_button

    add_notregistered_button = Button(notregistered_item_frame,text="Ekle",width = 10,command=addNotRegisteredItem)
    back_button = Button(notregistered_item_frame,text="Geri",width = 10,command=backToRegister)
    clear_button = Button(notregistered_item_frame,text="Temizle",width = 10,command=clear)
    increase_button = Button(notregistered_item_frame,text="Seçili Ürünü Arttır",width = 15,command=increase)


    editBack_item_button = Button(edit_frame,text ='Geri',width = 10, height =1,bg="light steel blue",command = lambda: back(add_button,remove_button,edit_button,add_frame,remove_frame,edit_frame))
    
    ## Deleting and adding registered product to registered list from inventory.
    global notregistered_removeitem_button
    global registered_removeitem_button

    notregistered_removeitem_button = Button(remove_frame,text='Ürün Sil',width=15,height=2,bg="gray",command = removeNotRegistered)
    registered_removeitem_button = Button(inventory_frame2,text="Depodaki Ürünü Zimmetle",width=20,height=2,bg='light steel blue',command=removeRegistered)
    registered_removeitem_button.place(x=120,y=150)

    global notregistered_remove_frame
    global registered_remove_frame

    notregistered_remove_frame = Frame(remove_frame,width=400,height=442,highlightbackground='black',highlightthickness=1,bg="LavenderBlush2")
    registered_remove_frame = Frame(remove_frame,width=400,height=442,highlightbackground='black',highlightthickness=1,bg="LavenderBlush2")

    global malzemey,miktary,aciklamay,miktary2,sec,ekley,geriy,temizley,sec2,vereny,alany,tarihy,vereny2,alany2,tarihy2,ekley2
    global smalzeme,smalzeme_metin,solcu,saciklama

    smalzeme = Label(registered_remove_frame,text="",bg="LavenderBlush2")
    smalzeme_metin = Label(registered_remove_frame,text="",bg="LavenderBlush2")
    solcu = Label(registered_remove_frame,text="",bg="LavenderBlush2")
    

    malzemey = Entry(registered_remove_frame)

    miktary = Entry(registered_remove_frame,validate="key")
    miktary['validatecommand'] = (miktary.register(testVal),'%P','%d')

    aciklamay = Entry(registered_remove_frame)

    vereny = Entry(registered_remove_frame)
    alany = Entry(registered_remove_frame)
    tarihy = Entry(registered_remove_frame)

    miktary2 = Entry(registered_remove_frame,validate="key")
    miktary2['validatecommand'] = (miktary2.register(testVal),'%P','%d')
    saciklama = Entry(registered_remove_frame)
    vereny2 = Entry(registered_remove_frame)
    alany2 = Entry(registered_remove_frame)
    tarihy2 = Entry(registered_remove_frame)
    

    sec = Button(registered_remove_frame,text="Seç",width=10,command=selectRegistered)
    ekley = Button(registered_remove_frame,text="Ekle",width=10,command=noRegister)
    ekley2 = Button(registered_remove_frame,text="Ekle",width=10,command=selectRegister)
    geriy = Button(registered_remove_frame,text="Geri",width=10,command=backagain)
    temizley = Button(registered_remove_frame,text="Temizle",width=10,command=clearagain)
    


    global malzeme_no, sayi, sayi2, hurdaaciklama,hurdaaciklama2

    malzeme_no = Entry(notregistered_remove_frame)
    
    
    sayi = Entry(notregistered_remove_frame,validate="key")
    sayi['validatecommand'] = (sayi.register(testVal),'%P','%d')

    hurdaaciklama = Entry(notregistered_remove_frame)
    sayi2 = Entry(notregistered_remove_frame,validate="key")
    sayi2['validatecommand'] = (sayi2.register(testVal),'%P','%d')

    hurdaaciklama2 = Entry(notregistered_remove_frame)

    global remove_notregistered, remove_notregistered2, geri,clear_remove2

    remove_notregistered = Button(notregistered_remove_frame,text="Çıkar",width=10,command=removeNotRegisteredItem)
    remove_notregistered2 = Button(notregistered_remove_frame,text="Seçili Ürünü Çıkar",width=15,command=decrease)
    geri = Button(notregistered_remove_frame,text="Geri",width=10,command=backRemoveRegister)
    clear_remove2 = Button(notregistered_remove_frame,text="Temizle",width=10,command=clear_remove)
    

    ## editing nonregistered product

    global notregistered_edit_frame
    global registered_edit_frame

    global notregistered_edit_button
    global registered_edit_button

    notregistered_edit_frame =Frame(edit_frame,width=400,height=440,bg="LavenderBlush2")
    registered_edit_frame = Frame(edit_frame,width=400,height=440,bg="LavenderBlush2")

    notregistered_edit_button = Button(edit_frame,text="Malzeme No'ya Göre",width=15,height=2,bg="light steel blue",command=no)
    registered_edit_button = Button(edit_frame,text="Seçilen Ürüne Göre",width=15,height=2,bg="light steel blue",command=selectedEdit)

    global malzeme_no1,malzeme_no2,malzeme_metin2,miktar2,olcu2,aciklama2
    malzeme_no1 = Entry(notregistered_edit_frame)
    malzeme_no2 = Entry(notregistered_edit_frame)
    malzeme_metin2 = Entry(notregistered_edit_frame)

    miktar2 = Entry(notregistered_edit_frame,validate="key")
    miktar2['validatecommand'] = (miktar2.register(testVal),'%P','%d')

    olcu2 = Entry(notregistered_edit_frame)
    aciklama2 = Entry(notregistered_edit_frame)

    global searchB,editB,backB,clearB

    searchB = Button(notregistered_edit_frame,text="Bul",width=10,command=search)
    editB = Button(notregistered_edit_frame,text="Düzenle",width=10,command=edit1)
    editB['state'] = DISABLED
    backB = Button(notregistered_edit_frame,text="Geri",width=10,command=backToEdit)
    clearB = Button(notregistered_edit_frame,text="Temizle",width=10,command=clearEdit)


    global malzeme_no3,malzeme_metin3,miktar3,olcu3,aciklama3
    malzeme_no3 = Entry(registered_edit_frame)
    malzeme_metin3 = Entry(registered_edit_frame)

    miktar3 = Entry(registered_edit_frame,validate="key")
    miktar3['validatecommand'] = (miktar3.register(testVal),'%P','%d')


    olcu3 = Entry(registered_edit_frame)
    aciklama3 = Entry(registered_edit_frame)

    global searchB2,editB2,backB2,clearB2

    searchB2 = Button(registered_edit_frame,text="Seç",width=10,command=findSelected)
    editB2 = Button(registered_edit_frame,text="Düzenle",width=10,command=edit2)
    editB2['state'] = DISABLED
    backB2 = Button(registered_edit_frame,text="Geri",width=10,command=backEdit)
    clearB2 = Button(registered_edit_frame,text="Temizle",width=10,command=clearEdit2)


    ############################################################################################################
    

    

    
    ################################################# Registered product List
    global register_select
    register_select = Button(window,text = "Zimmet Listesi",width=11,bg="white",command=lambda:showRegister(inventory_frame,inventory_frame2,register_frame,register_frame2))
    register_select.place(x=1150,y=30)
    
    #frames of registered product list
    register_frame = Frame(window,width=1050,height=600,highlightbackground='black',highlightthickness=3)
    register_frame2 = Frame(window,highlightbackground='black',width=250,height=450,bg ="lavender")

    global remove_product,edit_product,back_product
    #buttons of registered product list
    remove_product = Button(register_frame2,text="Ürün Sil",width=15,height=2,bg = 'light steel blue',command=removeRegister)
    back_product = Button(register_frame2,text="Ürünü Depoya Geri Al",width=18,height=2,bg = 'light steel blue',command=backInventory)
    edit_product = Button(register_frame2,text="Ürün Düzenle",width=15,height=2,bg = 'light steel blue',command=editRegister)

    # Edit part of registered product list
    global malzB,secB,geriB

    global remove_register_frame,edit_register_frame,back_register_frame
    remove_register_frame = Frame(register_frame2,width=245,height=445,highlightbackground='black',highlightthickness=1,bg="LavenderBlush2")
    edit_register_frame = Frame(register_frame2,width=245,height=445,highlightbackground='black',highlightthickness=1,bg="LavenderBlush2")
    back_register_frame = Frame(register_frame2,width=245,height=445,highlightbackground='black',highlightthickness=1,bg="LavenderBlush2")

    global malzNo_frame, sec_Frame

    malzNo_frame = Frame(edit_register_frame,width=243,height=443,bg="LavenderBlush2")
    sec_Frame = Frame(edit_register_frame,width=243,height=443,bg="LavenderBlush2")
    
    
    malzB = Button(edit_register_frame,text="Malzeme No'ya Göre",width=18,height=2,bg = 'light steel blue',command=malzNo)
    secB = Button(edit_register_frame,text="Seçilen Ürüne Göre",width=18,height=2,bg = 'light steel blue',command=sec_fr)
    geriB = Button(edit_register_frame,text="Geri",width=15,height=1,bg = 'light steel blue',command=backRegis)


    ## search product id to edit registered product.
    global malzemeq,malzemeq2,malzeme_metinq,miktarq,olcuq,verenq,alanq,tarihq,aciklamaq
    global bulq,duzenleq,geriq,temizleq

    malzemeq = Entry(malzNo_frame)
    malzemeq2 = Entry(malzNo_frame)
    malzeme_metinq = Entry(malzNo_frame)

    miktarq = Entry(malzNo_frame,validate="key")
    miktarq['validatecommand'] = (miktarq.register(testVal),'%P','%d')
    
    olcuq = Entry(malzNo_frame)
    verenq = Entry(malzNo_frame)
    alanq = Entry(malzNo_frame)
    tarihq = Entry(malzNo_frame)
    aciklamaq = Entry(malzNo_frame)

    bulq = Button(malzNo_frame,text="Bul",width = 10,command=searchRegist)
    duzenleq = Button(malzNo_frame,text="Düzenle",width = 10,command=edit3)
    geriq = Button(malzNo_frame,text="Geri",width = 10,command=backEditto)
    temizleq = Button(malzNo_frame,text="Temizle",width = 10,command=clearEditto)

    duzenleq['state'] = DISABLED

    ## select product to edit registered product.
    global malzemet,malzemet2,malzeme_metint,miktart,olcut,verent,alant,tariht,aciklamat
    global sect,duzenlet,gerit,temizlet

    malzemet = Entry(sec_Frame)
    malzeme_metint = Entry(sec_Frame)

    miktart = Entry(sec_Frame,validate="key")
    miktart['validatecommand'] = (miktart.register(testVal),'%P','%d')

    olcut= Entry(sec_Frame)
    verent = Entry(sec_Frame)
    alant = Entry(sec_Frame)
    tariht = Entry(sec_Frame)
    aciklamat = Entry(sec_Frame)

    sect = Button(sec_Frame,text="Seç",width = 10,command=selectReg)
    duzenlet = Button(sec_Frame,text="Düzenle",width = 10,command=edit4)
    gerit = Button(sec_Frame,text="Geri",width = 10,command=backEditto2)
    temizlet = Button(sec_Frame,text="Temizle",width = 10,command=clearEditto2)
    
    duzenlet['state'] = DISABLED
    # export button
    global excel_button2
    excel_button2 = Button(register_frame,text = 'Excele Aktar',width=20,height=2,bg = 'cadet blue',command= lambda: export2("local"))
    excel_button2.place(x=650,y=545)
    
    # filtering options
    global control_menu
    control_menu = StringVar(register_frame)
    option = ("Ürün: ","Kişi: ")
    control_menu.set("Seç:")

    category_label2 = Label(register_frame,text="Kategori Seç: ")
    category_label2.place(x=45,y=541)
    option_menu = OptionMenu(register_frame,control_menu,*option)
    option_menu.config(bg="cadet blue")
    option_menu["menu"].config(bg="white")
    option_menu.place(x=45,y=560)

    global category_entry2,filter_button2,filter_cancel2
    category_entry2 = Entry(register_frame,width = 25)
    category_entry2.place(x=140,y=565)
    
    # filter system of registered product list
    filter_button2 = Button(register_frame,text = 'Filtrele',width=20,height=2,bg='cadet blue',command=filter2)
    filter_button2.place(x=350,y=545)
    filter_cancel2 = Button(register_frame,text = 'X',width=2,height=2,bg='salmon',command= returnTable2)
    filter_cancel2.place(x=520,y=545)
    filter_cancel2['state'] = DISABLED

    
    # removing registered product

    global malzemer,miktarr,miktarr2,cikarr,cikarr2,gerir,temizler,hurdaaciklama3,hurdaaciklama4

    malzemer = Entry(remove_register_frame)

    miktarr = Entry(remove_register_frame,validate="key")
    miktarr['validatecommand'] = (miktarr.register(testVal),'%P','%d')

    hurdaaciklama3 = Entry(remove_register_frame)

    miktarr2 = Entry(remove_register_frame,validate="key")
    miktarr2['validatecommand'] = (miktarr2.register(testVal),'%P','%d')

    hurdaaciklama4 = Entry(remove_register_frame)

    cikarr = Button(remove_register_frame,text="Çıkar",width = 10,command=deleteReg)
    cikarr2 = Button(remove_register_frame,text="Seçili Ürünü Çıkar",width = 15,command=deleteSelectedReg)
    gerir = Button(remove_register_frame,text="Geri",width = 10,command=backReg)
    temizler = Button(remove_register_frame,text="Temizle",width = 10,command=clearReg)

    global malzeme0,miktar0,aciklama0,veren0,alan0,miktar02,aciklama02
    global malzeme0x, malzeme_metin0x,olcu0x,veren0x,alan0x,tarih0x
    global ekle0,sec0,ekle02,geri02,temizle02

    malzeme0 = Entry(back_register_frame)

    miktar0 = Entry(back_register_frame,validate="key")
    miktar0['validatecommand'] = (miktar0.register(testVal),'%P','%d')
    
    aciklama0 = Entry(back_register_frame)
    veren0 = Entry(back_register_frame)
    alan0 = Entry(back_register_frame)

    miktar02 = Entry(back_register_frame,validate="key")
    miktar02['validatecommand'] = (miktar02.register(testVal),'%P','%d')

    aciklama02 = Entry(back_register_frame)

    # adding registered product back to inventory as nonregisterede product
    malzeme0x = Label(back_register_frame,text="",bg="LavenderBlush2")
    malzeme_metin0x = Label(back_register_frame,text="",bg="LavenderBlush2")
    olcu0x = Label(back_register_frame,text="",bg="LavenderBlush2")
    veren0x = Label(back_register_frame,text="",bg="LavenderBlush2")
    alan0x = Label(back_register_frame,text="",bg="LavenderBlush2")
    tarih0x = Label(back_register_frame,text="",bg="LavenderBlush2")

    ekle0 = Button(back_register_frame,text="Ekle",width = 10,command=addBack)
    sec0 = Button(back_register_frame,text="Seç",width = 10,command=selecting)
    ekle02 = Button(back_register_frame,text="Ekle",width = 8,command=addBack2)
    geri02 = Button(back_register_frame,text="Geri",width = 8,command = backMain)
    temizle02 = Button(back_register_frame,text="Temizle",width = 8,command=clearAll)

    ekle02['state'] = DISABLED
       
    ## JUNK LIST
    global junk_select
    junk_select = Button(window,text= "Hurda Listesi",width = 10,bg = "white",command=lambda : showJunk(inventory_frame,inventory_frame2,register_frame,register_frame2))
    junk_select.place(x=1245,y=30)
    
    global junk_frame
    
    junk_frame = Frame(window,width=1050,height=600,highlightbackground='black',highlightthickness=3)

    global category_entry3,filter_button3,filter_cancel3

    category_label3 = Label(junk_frame,text = 'Ürün İsmi Giriniz: ')
    category_label3.place(x=30,y=555)
    category_entry3 = Entry(junk_frame,width = 25)
    category_entry3.place(x=140,y=557)
    
    filter_button3 = Button(junk_frame,text = 'Filtrele',width=20,height=2,bg='cadet blue',command=filter3)
    filter_button3.place(x=350,y=545)
    filter_cancel3 = Button(junk_frame,text = 'X',width=2,height=2,bg='salmon',command = returnTable3)
    filter_cancel3.place(x=520,y=545)

    global excel_button3
    excel_button3 = Button(junk_frame,text = 'Excele Aktar',width=20,height=2,bg="cadet blue",command=lambda: export3("local"))
    excel_button3.place(x=650,y=545)

    
    filter_cancel3['state'] = DISABLED

    global clearallbutton
    clearallbutton = Button(junk_frame,text = 'Listeyi Temizle',width=15,height=2,bg="salmon",command=removeAll)
    clearallbutton.place(x=910,y=545)

    

    ########################################################################### TABLE - DATA ######################################################

    # Table properties
    style = ttk.Style()
    style.theme_use("default")
    style.configure("Treeview", 
	background="azure",
	foreground="black",
	rowheight=51,
	fieldbackground="azure",
    font = (None,8)
	)
    style.map('Treeview', 
	background=[('selected', 'blue')])

    #creating three tables which are inventory list, registered product list and junk list.  
    global table,table2,table3

    table = ttk.Treeview(inventory_frame, selectmode="browse")
    table2 = ttk.Treeview(register_frame, selectmode="browse")
    table3 = ttk.Treeview(junk_frame, selectmode="browse")
    
    #scrollbars of tables
    scrollbar = ttk.Scrollbar(inventory_frame, orient="vertical", command=table.yview)
    scrollbar.place(x=930, y=0, height=540)

    scrollbar2 = ttk.Scrollbar(inventory_frame, orient="horizontal", command=table.xview)
    scrollbar2.place(x=0, y=524, width=930)

    scrollbar3 = ttk.Scrollbar(register_frame, orient="vertical", command=table2.yview)
    scrollbar3.place(x=1030, y=0, height=542)

    scrollbar4 = ttk.Scrollbar(register_frame, orient="horizontal", command=table2.xview)
    scrollbar4.place(x=0, y=527, width=1030)

    scrollbar5 = ttk.Scrollbar(junk_frame, orient="vertical", command=table3.yview)
    scrollbar5.place(x=1030, y=0, height=542)

    scrollbar6 = ttk.Scrollbar(junk_frame, orient="horizontal", command=table3.xview)
    scrollbar6.place(x=0, y=527, width=1030)

    table.configure(yscrollcommand=scrollbar.set,xscrollcommand=scrollbar2.set)
    table2.configure(yscrollcommand=scrollbar3.set,xscrollcommand=scrollbar4.set)
    table3.configure(yscrollcommand=scrollbar5.set,xscrollcommand=scrollbar6.set)

    table.place(x=0,y=0)
    table2.place(x=0,y=0)
    table3.place(x=0,y=0)

    #load data.
    importExcel()
    importExcel2()
    importExcel3()

    global p,p2,p3
    p = 0
    p2 = 0
    p3 = 0
    # saving before quiting program.
    window.protocol("WM_DELETE_WINDOW",beforeExit)
    window.mainloop()
    
main()


    


