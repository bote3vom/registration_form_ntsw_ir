
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client as win32
import os

# Persian fields in the Excel template
fields = [
    "تاریخ سند",
    "شماره صورتحساب",
    "کد/شناسه ملی خریدار",
    "نام خریدار",
    "تلفن همراه",
    "کدپستی انبار مبدا",
    "شرح سند",
    "شناسه کالا",
    "تعداد/مقدار",
    "مبلغ واحد (ریال)",
    "مبلغ تخفیف (ریال)",
    "سایر اضافات (ریال)",
    "مبلغ مالیات و عوارض (ریال)"
]

records = []

root = tk.Tk()
root.title("نرم افزار ثبت فروش")
root.geometry("950x750")

entries = {}

# Input form
form = ttk.LabelFrame(root, text="اطلاعات فاکتور")
form.pack(fill="x", padx=10, pady=10)

for i, field in enumerate(fields):
    row = i // 2
    col = (i % 2) * 2

    label = ttk.Label(form, text=field)
    label.grid(row=row, column=col, padx=8, pady=7)

    entry = ttk.Entry(form, width=27, justify="right")
    entry.grid(row=row, column=col + 1, padx=8, pady=7)

    entries[field] = entry


# Add a record
def add_record():

    data = [entries[field].get().strip() for field in fields]

    if not data[0] or not data[5] or not data[7]:
        messagebox.showerror(
            "خطا",
            "تاریخ سند، کدپستی انبار و شناسه کالا الزامی هستند."
        )
        return

    try:
        quantity = float(data[8])
        unit_price = int(data[9])

        discount = int(data[10] or 0)
        additions = int(data[11] or 0)
        tax = int(data[12] or 0)

    except ValueError:
        messagebox.showerror(
            "خطا",
            "مقادیر تعداد و مبالغ باید عددی باشند."
        )
        return

    data[8] = quantity
    data[9] = unit_price
    data[10] = discount
    data[11] = additions
    data[12] = tax

    records.append(data)

    table.insert("", "end", values=data)

    for entry in entries.values():
        entry.delete(0, tk.END)

    messagebox.showinfo("موفق", "اطلاعات ثبت شد.")


# Save records to XLSM
def save_excel():

    if not records:
        messagebox.showwarning("هشدار", "هیچ اطلاعاتی وارد نشده است.")
        return

    source = filedialog.askopenfilename(
        title="انتخاب فایل اکسل",
        filetypes=[("Excel Macro Enabled", "*.xlsm")]
    )

    if not source:
        return

    output = filedialog.asksaveasfilename(
        title="ذخیره فایل خروجی",
        defaultextension=".xlsm",
        filetypes=[("Excel Macro Enabled", "*.xlsm")],
        initialfile="Sales_Output.xlsm"
    )

    if not output:
        return

    excel = None
    wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        # Disable macro execution while opening the template.
        excel.AutomationSecurity = 3

        wb = excel.Workbooks.Open(
            os.path.abspath(source),
            UpdateLinks=0,
            ReadOnly=False
        )

        # Main worksheet
        ws = wb.Worksheets("فروش به مصرف کننده")

        # Save as a separate XLSM file
        wb.SaveAs(
            os.path.abspath(output),
            FileFormat=52
        )

        # Write records starting at row 2
        for i, record in enumerate(records, start=2):

            for j, value in enumerate(record, start=1):

                cell = ws.Cells(i, j)

                # Preserve IDs, dates and telephone numbers as text
                if j in [1, 2, 3, 5, 6, 8]:
                    cell.NumberFormat = "@"
                    cell.Value = str(value)
                else:
                    cell.Value = value

        wb.Save()
        wb.Close(SaveChanges=True)
        wb = None

        excel.Quit()
        excel = None

        messagebox.showinfo(
            "موفق",
            f"فایل با موفقیت ذخیره شد:\n{output}"
        )

    except Exception as e:

        messagebox.showerror(
            "خطا",
            str(e)
        )

        if wb is not None:
            wb.Close(SaveChanges=False)

        if excel is not None:
            excel.Quit()


# Buttons
buttons = ttk.Frame(root)
buttons.pack(pady=10)

ttk.Button(
    buttons,
    text="افزودن اطلاعات",
    command=add_record
).pack(side="right", padx=10)

ttk.Button(
    buttons,
    text="ذخیره در اکسل",
    command=save_excel
).pack(side="right", padx=10)


# Preview table
table_frame = ttk.LabelFrame(root, text="اطلاعات وارد شده")
table_frame.pack(fill="both", expand=True, padx=10, pady=10)

table = ttk.Treeview(
    table_frame,
    columns=fields,
    show="headings"
)

for field in fields:
    table.heading(field, text=field)
    table.column(field, width=130, anchor="center")

scroll_y = ttk.Scrollbar(
    table_frame,
    orient="vertical",
    command=table.yview
)

scroll_x = ttk.Scrollbar(
    table_frame,
    orient="horizontal",
    command=table.xview
)

table.configure(
    yscrollcommand=scroll_y.set,
    xscrollcommand=scroll_x.set
)

scroll_y.pack(side="right", fill="y")
scroll_x.pack(side="bottom", fill="x")
table.pack(fill="both", expand=True)

root.mainloop()
