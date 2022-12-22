import tkinter as tk
from tkinter import filedialog
from openpyxl import load_workbook, Workbook
from datetime import date


def get_stores():
    """
    Imports an xslx file containing all stores for a given region
    """
    global ws1
    file_path = filedialog.askopenfilename()
    wb1 = load_workbook(filename=file_path)
    ws1 = wb1[wb1.sheetnames[0]]
    if ws1:
        storeLabel = tk.Label(text="Butikslista importerad")
        storeLabel.place(x=160, y=10)
    
    
def get_visists():
    """
    Imports the excel file containing visits and their dates
    """
    global ws2
    file_path = filedialog.askopenfilename()
    wb2 = load_workbook(filename=file_path)
    ws2 = wb2[wb2.sheetnames[0]]
    if ws2:
        visitLabel = tk.Label(text="Besökslista importerad")
        visitLabel.place(x=160, y=40)


def export():
    """
    Work through the two imported worksheets to create a new excel file with store and days since you visisted. 
    """
    region_dict = {}
    
    # fyll en dictionary med {butik: {adress: , ort: , Tid: }}
    for i in ws1:
        region_dict.update({i[0].value: {"Adress": i[1].value, "Ort": i[2].value, "Tid": "N/A"}})
    
    # Nu ska vi ta fram tid från ws2
    for i in ws2:
        if i[0].value in region_dict:
            # date.today() - i[2].value.date() is a timedelta and needs to be turned into a date?
            if region_dict[i[0].value]["Tid"] == "N/A" or region_dict[i[0].value]["Tid"] > (date.today() - i[2].value.date()).days:
                # This seems to become hours when I save it to my file...
                region_dict[i[0].value]["Tid"] = (date.today() - i[2].value.date()).days
    
    
    # skapa en ny workbook och ett nytt sheet
    export = Workbook()
    export_sheet = export.active
    
    # worksheet.append([list of items ot append]) seems to be the way to do this
    # butiksnamn, adress, ort, dagar. Note: vi itererar genom nycklarna för vår yttre dictionary, så store är enbart en string. 
    # Gör region_dict[store]["Tid"] om du vill ha tiden.
    for store in region_dict:
        export_sheet.append([store, region_dict[store]["Adress"], region_dict[store]["Ort"], region_dict[store]["Tid"]])
    
    export.save("test_file.xlsx")
    

root = tk.Tk()
root.title("Dagar sedan besök")
root.geometry("500x200+500+500")
root.configure(bg='lightblue')


importStores = tk.Button(root, text="Importera Regionslista", command=get_stores, bg="lightgrey")
importStores.place(x=10, y=10)
importStores.configure(border=2, relief="raised")

importVisits = tk.Button(root, text="Importera besökslista  ", command=get_visists, bg="lightgrey")
importVisits.place(x=10, y=40)
importVisits.configure(border=2, relief="raised")

exportButton = tk.Button(root, text="Exportera besökslista  ", command=export,  bg="orange")
exportButton.place(x=10, y=160)
exportButton.configure(border=2, relief="raised")

quitButton = tk.Button(root, text="Avsluta", command=root.quit)
quitButton.place(x=430, y=160)
quitButton.configure(border=2, relief="raised")

root.mainloop()


