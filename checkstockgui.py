import tkinter as tk
from datetime import datetime
from tkinter import filedialog,ttk
import random
from openpyxl import load_workbook,Workbook

export = []

def openexcel():
    global max_row,sheet
    file_path = filedialog.askopenfilename()
    if file_path:
        result_openfile.configure(text=f'เปิดไฟล์: {file_path}')
    file = load_workbook(file_path)
    sheet = file.worksheets[0]
    max_row = sheet.max_row

def exportexcel():
    if len(export) > 0:
        crr_d = datetime.now()
        day = crr_d.day
        month = crr_d.month
        year = crr_d.year
        yearMonthDay = f'{year}-{month}-{day}'
        excel = Workbook()
        ran = random.randrange(1,1000)
        excel_name = f"stock-export{yearMonthDay}{ran}.xlsx"
        excel.save(excel_name)
        work_sheet = excel.worksheets[0]
        work_sheet.title = 'สินค้าคงคลัง'
        header = ['SKU','ONHAND','OUTGOING','BALANCE']
        r = 2
        for items in export:
            for i in range(4):
                if (r-1) == 1:
                    work_sheet.cell(row=1, column=i+1).value = header[i]
                    work_sheet.cell(row=r, column=i+1).value = items[i]
                else:
                    work_sheet.cell(row=r, column=i+1).value = items[i]
            r += 1
        excel.save(excel_name)  
        result_label.configure(text=f'Export สำเร็จ: {excel_name}')
    else:result_label.configure(text=f'ไม่พบข้อมูลที่สามารถ export ได้')

def search_sku():
    try:
        col_onhand = int(onhand_entry.get())
        col_outgoing = int(outgoing_entry.get())
        col_sku = int(col_sku_entry.get())
    except ValueError:
        result_label.configure(text=f'คอลัมน์ไม่ถูกต้อง')
    search = sku_entry.get()
    s_sku = search.lower().replace('#','')
    is_found = False
    for num_row in range(max_row):
        row = num_row + 1
        sku = str(sheet.cell(row=row, column=col_sku).value).lower().replace('#','')
        if s_sku == sku:
            onhand = int(sheet.cell(row=row, column=col_onhand).value) or 0
            outgoing = int(sheet.cell(row=row, column=col_outgoing).value) or 0
            result = onhand - outgoing
            result_label.configure(text=f'{s_sku.upper()} สินค้าคงเหลือ: {result}')
            items = [s_sku.upper(), onhand, outgoing, result]
            if items not in export:
                export.append(items)
            is_found = True
            break
    if not is_found:
        result_label.configure(text=f'ไม่พบสินค้า: {s_sku}')
        
def get_sku_and_search(eent=None):
    search_sku()

root = tk.Tk()
root.geometry('600x800')
root.title('Search SKU')
root.iconbitmap('icon/favicon.ico')
root.option_add('*Font','Angsana_New 14')


howto_text = tk.StringVar()
howto_text.set('วิธีการใช้\n1. เปิดไฟล์ที่ต้องการด้วยปุ่ม Open file\n2. คอลัมน์ onhand ใส่เลขคอลัมน์ของปริมาณในมือ\n3. คอลัมน์ outgoing ใส่เลขคอลัมน์ของกล่องขาออก\n4. กรอกรหัสสินค้าที่ช่อง SKU กด Enter\n5. Export file จะส่งออกไฟล์ excel ตามข้อมูลสินค้าที่ค้นพบ')
howto = tk.Label(root,textvariable=howto_text,justify='left',pady=20,height=8)
howto.pack()


button_frame = ttk.Frame(root)
button_frame.pack(pady=10)


openfile_button = ttk.Button(button_frame, text='Open file', command=openexcel)
openfile_button.pack(side='left')

export_button = ttk.Button(button_frame, text='Export file', command=exportexcel)
export_button.pack(side='left',padx=10)

result_openfile = tk.Label(root, text='', width=50, height=2)
result_openfile.pack(pady=5,ipady=10)

col_sku_label = tk.Label(root, text='คอลัมน์ SKU:')
col_sku_label.pack(pady=5)
col_sku_entry = ttk.Entry(root)
col_sku_entry.pack()

onhand_label = tk.Label(root, text='คอลัมน์ onhand:')
onhand_label.pack(pady=5)

onhand_entry = ttk.Entry(root)
onhand_entry.pack()

outgoing_label = ttk.Label(root, text='คอลัมน์ outgoing:')
outgoing_label.pack(pady=5)

outgoing_entry = ttk.Entry(root)
outgoing_entry.pack()

sku_label = tk.Label(root, text='SKU:')
sku_label.pack(pady=5)

sku_entry = ttk.Entry(root)
sku_entry.pack()

search_button = ttk.Button(root, text='Search', command=search_sku)
search_button.pack(pady=10)
sku_entry.bind('<Return>',get_sku_and_search)

result_label = tk.Label(root, text='', width=50, height=2)
result_label.pack(pady=20,ipady=10)

root.mainloop()
