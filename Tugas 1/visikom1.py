import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd


# Fungsi untuk menyimpan data intensitas nilai ke dalam bentuk excel
def simpan_nilai_intensitas_to_excel(image):
    data = []
    baris, kolom, _ = image.shape
    for i in range(baris):
        for j in range(kolom):
            R, G, B = image[i, j]
            data.append([f"f({i}, {j})", R, G, B])
    
    df = pd.DataFrame(data, columns=["Pixel", "R", "G", "B"])

    max_sheet_size = 1048576  # Maksimum ukuran lembar kerja Excel
    num_rows = df.shape[0]
    num_sheets = (num_rows - 1) // max_sheet_size + 1

    with pd.ExcelWriter("nilai_pixel.xlsx") as writer:
        for i in range(num_sheets):
            start_idx = i * max_sheet_size
            end_idx = min((i + 1) * max_sheet_size, num_rows)
            df_subset = df.iloc[start_idx:end_idx]
            sheet_name = f"Sheet_{i+1}"
            df_subset.to_excel(writer, sheet_name=sheet_name, index=False)


# Fungsi untuk membuat Rotasi dan Flip (Pencerminan)
def rotasi_90_derajat_kanan(matriks):
    baris, kolom, _ = matriks.shape
    matriks_rotasi = matriks.copy() # matriks.copy (membuat salinan dari gambar asli)
    for i in range(baris):
        for j in range(kolom):
            matriks_rotasi[j, baris-1-i] = matriks[i, j]
    return matriks_rotasi

def rotasi_90_derajat_kiri(matriks):
    baris, kolom, _ = matriks.shape
    matriks_rotasi = matriks.copy()
    for i in range(baris):
        for j in range(kolom):
            matriks_rotasi[kolom-1-j, i] = matriks[i, j]
    return matriks_rotasi

def flip_horizontal(matriks):
    baris, kolom, _ = matriks.shape
    matriks_pencerminan = matriks.copy()
    for i in range(baris):
        for j in range(kolom):
            matriks_pencerminan[baris-1-i, j] = matriks[i, j]
    return matriks_pencerminan

def flip_vertikal(matriks):
    baris, kolom, _ = matriks.shape
    matriks_pencerminan = matriks.copy()
    for i in range(baris):
        for j in range(kolom):
            matriks_pencerminan[i, kolom-1-j] = matriks[i, j]
    return matriks_pencerminan

# Fungsi untuk memuat gambar
def load_image():
    file_path = filedialog.askopenfilename()
    if file_path:
        try:
            image = cv2.imread(file_path)
            if image is not None:
                return image
            else:
                messagebox.showerror("Error", "Failed to load image. Please choose the correct image.")
        except Exception as e:
            messagebox.showerror("Error", f"Error occurred while loading image: {str(e)}")
    return None

# Fungsi untuk memproses gambar
def process_image():
    image = load_image()
    if image is not None:
        cv2.imshow("Original Image", image)
        cv2.waitKey(0)

        # Simpan nilai intensitas gambar ke dalam file Excel
        simpan_nilai_intensitas_to_excel(image)

        # Rotasi 90 derajat kearah kanan (searah jarum jam)
        rotasi_kanan = rotasi_90_derajat_kanan(image)
        cv2.imshow("Rotasi 90 Derajat Searah Jarum Jam", rotasi_kanan)
        cv2.waitKey(0)

        # Rotasi 90 derajat kearah kiri (berlawanan arah jarum jam)
        rotasi_kiri = rotasi_90_derajat_kiri(image)
        cv2.imshow("Rotasi 90 Derajat Berlawanan Arah Jarum Jam", rotasi_kiri)
        cv2.waitKey(0)

        # Pencerminan horizontal
        pencerminan_horizontal = flip_horizontal(image)
        cv2.imshow("Pencerminan Horizontal", pencerminan_horizontal)
        cv2.waitKey(0)

        # Pencerminan vertikal
        pencerminan_vertical = flip_vertikal(image)
        cv2.imshow("Pencerminan Vertical", pencerminan_vertical)
        cv2.waitKey(0)

        cv2.destroyAllWindows()

# Create GUI window
window = tk.Tk()
window.title("Image Processor")

# Create buttons
load_button = tk.Button(window, text="Choose Image", command=process_image)
load_button.pack(pady=10)

# Run the GUI loop
window.mainloop()
