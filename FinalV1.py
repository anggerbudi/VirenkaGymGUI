# Program GUI Pendaftaran Virenka Gym
# Coder - Martinus Angger Budi Wicaksono
# Version : 1.0.0
# Creation date : 12/02/23


# Mengimport library
import tkinter as tk
import customtkinter as ctk
from tkcalendar import DateEntry
from tkinter import messagebox
from PIL import Image
import openpyxl
import os.path
from openpyxl import Workbook

class VirenkaGym:
    def __init__(self, master):
        self.master = master
        master.title("Form Pendaftaran Virenka Gym")
        master.resizable(False, False)
        master.geometry("1280x720")

        # set background image
        my_image = ctk.CTkImage(light_image=Image.open("12066157_4883941.jpg"),
                                size=(1280, 720))
        background = ctk.CTkLabel(self.master, image=my_image, text="")
        background.place(x=0, y=0, relwidth=1, relheight=1)

        # Label dan entry
        self.nama_label = tk.Label(self.master, text="Nama:")
        self.nama_entry = tk.Entry(self.master)
        self.tempat_lahir_label = tk.Label(self.master, text="Tempat Lahir:")
        self.tempat_lahir_entry = tk.Entry(self.master)
        self.tanggal_lahir_label = tk.Label(self.master, text="Tanggal Lahir:")
        self.tanggal_lahir_entry = DateEntry(self.master, date_pattern='DD/MM/YYYY')
        self.alamat_label = tk.Label(self.master, text="Alamat:")
        self.alamat_entry = tk.Entry(self.master)
        self.nomor_telepon_label = tk.Label(self.master, text="Nomor Telepon:")
        self.nomor_telepon_entry = tk.Entry(self.master)
        self.jenis_kelamin_label = tk.Label(self.master, text="Jenis Kelamin:")
        self.agama_label = tk.Label(self.master, text="Agama:")

        # Radiobutton jenis kelamin
        self.jenis_kelamin_var = tk.StringVar()
        self.laki_laki_radiobutton = tk.Radiobutton(self.master, text="Laki-laki", variable=self.jenis_kelamin_var,
                                                    value="Laki-laki")
        self.perempuan_radiobutton = tk.Radiobutton(self.master, text="Perempuan", variable=self.jenis_kelamin_var,
                                                    value="Perempuan")

        # Default pilihan radiobutton jenis kelamin
        self.laki_laki_radiobutton.select()

        # Radio button agama
        self.agama_var = tk.StringVar()
        self.islam_radiobutton = tk.Radiobutton(self.master, text="Islam", variable=self.agama_var, value="Islam")
        self.kristen_radiobutton = tk.Radiobutton(self.master, text="Kristen", variable=self.agama_var, value="Kristen")
        self.katolik_radiobutton = tk.Radiobutton(self.master, text="Katolik", variable=self.agama_var, value="Katolik")
        self.hindu_radiobutton = tk.Radiobutton(self.master, text="Hindu", variable=self.agama_var, value="Hindu")
        self.buddha_radiobutton = tk.Radiobutton(self.master, text="Buddha", variable=self.agama_var, value="Buddha")
        self.konghucu_radiobutton = tk.Radiobutton(self.master, text="Konghucu", variable=self.agama_var, value="Konghucu")

        # Default pilihan radiobutton agama
        self.islam_radiobutton.select()

        # Submit button
        submit_button = tk.Button(self.master, text="Submit", command=self.submit)

        # Peletakan widgets pada grid
        self.nama_label.grid(row=0, column=0, padx=10, pady=10)
        self.nama_entry.grid(row=0, column=1, padx=10, pady=10)
        self.tempat_lahir_label.grid(row=1, column=0, padx=10, pady=10)
        self.tempat_lahir_entry.grid(row=1, column=1, padx=10, pady=10)
        self.tanggal_lahir_label.grid(row=2, column=0, padx=10, pady=10)
        self.tanggal_lahir_entry.grid(row=2, column=1, padx=10, pady=10)
        self.alamat_label.grid(row=3, column=0, padx=10, pady=10)
        self.alamat_entry.grid(row=3, column=1, padx=10, pady=10)
        self.nomor_telepon_label.grid(row=4, column=0, padx=10, pady=10)
        self.nomor_telepon_entry.grid(row=4, column=1, padx=10, pady=10)
        self.jenis_kelamin_label.grid(row=5, column=0, padx=10, pady=10)
        self.agama_label.grid(row=6, column=0, padx=10, pady=10)
        self.laki_laki_radiobutton.grid(row=5, column=1, padx=10, pady=10)
        self.perempuan_radiobutton.grid(row=5, column=2, padx=10, pady=10)
        self.islam_radiobutton.grid(row=6, column=1, padx=10, pady=10)
        self.kristen_radiobutton.grid(row=6, column=2, padx=10, pady=10)
        self.katolik_radiobutton.grid(row=7, column=1, padx=10, pady=10)
        self.hindu_radiobutton.grid(row=7, column=2, padx=10, pady=10)
        self.buddha_radiobutton.grid(row=8, column=1, padx=10, pady=10)
        self.konghucu_radiobutton.grid(row=8, column=2, padx=10, pady=10)
        submit_button.grid(row=9, column=1, padx=10, pady=10)

        # Mengecek file excel pada direktori aktif
        self.workbook_path = "VirenkaGym.xlsx"
        if os.path.isfile(self.workbook_path):
            self.workbook = openpyxl.load_workbook(self.workbook_path)
        else:
            self.workbook = Workbook()

    def submit(self):
        # Mengambil data dari entry
        nama = self.nama_entry.get()
        tempat_lahir = self.tempat_lahir_entry.get()
        tanggal_lahir = self.tanggal_lahir_entry.get()
        alamat = self.alamat_entry.get()
        nomor_telepon = self.nomor_telepon_entry.get()
        jenis_kelamin = self.jenis_kelamin_var.get()
        agama = self.agama_var.get()

        # Mengecek input pengguna
        if not nama or not tempat_lahir or not tanggal_lahir or not alamat or not nomor_telepon or not jenis_kelamin or not agama:
            messagebox.showerror("Error", "Harap mengisi semua kolom")
        else:
            # Pesan Sukses
            messagebox.showinfo(" Success", "Data berhasil disimpan.")
            messagebox.showinfo("Success", "Terima kasih telah mendaftar di Virenka Gym, {}!".format(nama))

            # Menuliskan data ke worksheet
            worksheet = self.workbook.active
            worksheet.append([nama, tempat_lahir, tanggal_lahir, alamat, nomor_telepon, jenis_kelamin, agama])

            # Menyimpan workbook
            self.workbook.save(self.workbook_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = VirenkaGym(root)
    root.mainloop()
