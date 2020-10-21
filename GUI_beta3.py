import tkinter as tk
import tkinter.font as tkFont
from PIL import ImageTk, Image
from subprocess import call

pat_id = ''

class ButtonEntry:
    def __init__(self, root):
        self.patientid_var = ""

        patientid_label = tk.Label(root, text="  Enter Patient ID    ")
        canvas_main.create_window(200, 460, window=patientid_label)

        self.patientid_entry = tk.Entry(root)
        canvas_main.create_window(400, 460, window=self.patientid_entry)

        doa_label = tk.Label(root, text="Enter Date of Admission ")
        canvas_main.create_window(200, 500, window=doa_label)

        self.doa_entry = tk.Entry(root)
        canvas_main.create_window(400, 500, window=self.doa_entry)

        plotButton = tk.Button(root, text="Get Patient Report", command=self.get_patient_report)
        canvas_main.create_window(300, 560, window=plotButton)

    def get_patient_report(self):
        global pat_id, pat_name

        self.patientid_var = self.patientid_entry.get()
        patid_label2 = tk.Label(root, text='Patient ID: ' + str(self.patientid_var))
        canvas_main.create_window(300, 600, window=patid_label2)

        self.doa_var = self.doa_entry.get()
        doa_label2 = tk.Label(root, text=str('DOA: ' + self.doa_var))
        canvas_main.create_window(300, 620, window=doa_label2)

        pat_id = self.patientid_var
        pat_doa = self.doa_var

        patient_details = str(pat_id) + ' ' + str(pat_doa)

        open_file_write = open('pat_id_temp.txt', 'w+')
        open_file_write.writelines(patient_details)
        open_file_write.close()

        call(['python', 'patientdischargerv2.0beta3.py'])

        status_done_label = tk.Label(root, text='Patient report created successfully!', bg='white', fg='green')
        canvas_main.create_window(300, 640, window=status_done_label)
        return


if __name__ == "__main__":
    root = tk.Tk()
    root.title('Patient Discharger v2.0')
    canvas_main = tk.Canvas(root, width=600, height=660)
    canvas_main.pack()

    fontStyle = tkFont.Font(family="Calibri", size=30)

    path = 'logo.gif'
    img = ImageTk.PhotoImage(Image.open(path))
    logo_panel = tk.Label(root, image=img)
    canvas_main.create_window(300, 200, window=logo_panel)

    departmentname_label = tk.Label(root, text="Department of General Medicine", font=fontStyle)
    canvas_main.create_window(300, 400, window=departmentname_label)

    BE = ButtonEntry(root)

    statusbar = tk.Label(root,
                         text="""Homegrown with """ + '\u2665\uFE0F' + """ \n by Aditya Kshetrapal (MBBS 2014) and Nandan Patel (MBBS 2019) """,
                         bd=0, relief=tk.SUNKEN, anchor=tk.S)
    statusbar.pack(side=tk.BOTTOM, fill=tk.X)

    root.mainloop()
