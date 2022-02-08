import tkinter.filedialog as filedialog
import tkinter as tk

from docxtpl import DocxTemplate
import os
import xlrd

master = tk.Tk()

def input():
    input_path = tk.filedialog.askopenfilename()
    input_entry.delete(1, tk.END)  # Remove current text in entry
    input_entry.insert(0, input_path)  # Insert the 'path'

def template():
    template_path = tk.filedialog.askopenfilename()
    template_entry.delete(1, tk.END)  # Remove current text in entry
    template_entry.insert(0, template_path)  # Insert the 'path'

def output():
    path = tk.filedialog.askdirectory()
    output_entry.delete(1, tk.END)  # Remove current text in entry
    output_entry.insert(0, path)  # Insert the 'path'

def data_task(my_excel, template_path, output_directory):
    print("Operation initiated !")

    ################### READ EXCEL VALUES #################
    excel_path = my_excel

    wb = xlrd.open_workbook(excel_path)
    sheet = wb.sheet_by_index(0)

    file_list = []

    for m in range(sheet.ncols):
        A_list = []
        for i in range(sheet.nrows -2): #-2 cause we have 2 headers for our table
            A_list.append(str(sheet.cell_value(i+2, m)))
        file_list.append(A_list)
    print(file_list)
    ################### READ EXCEL VALUES #################
    for i in range(len(file_list[0])):
        data1 = {}

        for m in range(len(file_list)):
            data1[f"A{m}"] = str(file_list[m][i])


        def get_context(data):
            # """ You can generate your context separately since you may deal with a lot 
            #     of documents. You can carry out computations, etc in here and make the
            #     context look like the sample below.
            # """
            return data

        template_file = template_path

        name_prefix = str(file_list[0][i]).rstrip("\n")

        target_file = output_directory + f"/{name_prefix}-Attestation_model.docx"

        def from_template(template, target_file):
            target_file = target_file

            template = DocxTemplate(template)
            context = get_context(data1)  # gets the context used to render the document

            template.render(context)
            template.save(target_file)

            return target_file

        from_template(template_file, target_file)
    print("Operation Ended !")

def begin():
    print("started")
    my_excel = input_entry.get()
    template_doc = template_entry.get()
    output_directory = output_entry.get()
    data_task(my_excel, template_doc, output_directory)


top_frame = tk.Frame(master)
middle_frame = tk.Frame(master)
bottom_frame = tk.Frame(master)
line = tk.Frame(master, height=1, width=400, bg="grey80", relief='groove')

input_path = tk.Label(top_frame, text="Input Excel File Path:")
input_entry = tk.Entry(top_frame, text="", width=40)
browse1 = tk.Button(top_frame, text="Browse", command=input)

template_path = tk.Label(middle_frame, text="Input Docx Template File Path:")
template_entry = tk.Entry(middle_frame, text="", width=40)
browse3 = tk.Button(middle_frame, text="Browse", command=template)

output_path = tk.Label(bottom_frame, text="Output Folder Path:")
output_entry = tk.Entry(bottom_frame, text="", width=40)
browse2 = tk.Button(bottom_frame, text="Browse", command=output)

begin_button = tk.Button(bottom_frame, text='Begin!', command=begin)

top_frame.pack(side=tk.TOP)
middle_frame.pack(side=tk.TOP)
line.pack(pady=10)
bottom_frame.pack(side=tk.BOTTOM)

input_path.pack(pady=5)
input_entry.pack(pady=5)
browse1.pack(pady=5)

template_path.pack(pady=5)
template_entry.pack(pady=5)
browse3.pack(pady=5)

output_path.pack(pady=5)
output_entry.pack(pady=5)
browse2.pack(pady=5)

begin_button.pack(pady=20, fill=tk.X)

##get cwd path
__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))
ico_path = os.path.join(__location__, 'devcorp-logo.ico')

try:
    master.iconbitmap(ico_path)
except:
    pass
master.title(string='Devcorp | Model Generator')

master.mainloop()

