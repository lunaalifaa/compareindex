import subprocess
import sys
import pandas as pd
import os
import openpyxl
import re

def install_package(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])

# Install dependencies if not available
try:
    import pandas as pd
    import os
    import openpyxl
    import re
except ImportError:
    install_package("pip")
    install_package("pandas")
    install_package("openpyxl")
    import pandas as pd
    import os
    import openpyxl
    import re

import os
import pandas as pd
import glob

# Tentukan folder tempat file Excel berada
folder_path = "c:/Users/hp/Downloads/Documents/New folder/"

# Ambil semua file .xlsx dalam folder
excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

# Gabungkan semua file ke dalam satu DataFrame
df_list = []

for file in excel_files:
    df = pd.read_excel(file, header=9)  # Pastikan header ada di baris ke-9
    df["Source File"] = os.path.basename(file)  # Tambahkan nama file asal
    df_list.append(df)

# Gabungkan semua data
df_index_member = pd.concat(df_list, ignore_index=True)

# Cek apakah kolom "Index" sudah ada
print(df_index_member.columns)

# Tentukan folder tempat file Excel berada
folder_path = "c:/Users/hp/Downloads/Documents/New folder/"

# Ambil semua file .xlsx dalam folder
excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

# Gabungkan semua file ke dalam satu DataFrame
df_list = []  # List untuk menyimpan DataFrame dari setiap file

for file in excel_files:
    df = pd.read_excel(file)  # Membaca file Excel
    df["Source File"] = os.path.basename(file)  # Tambahkan kolom nama file asal
    df_list.append(df)

# Gabungkan semua DataFrame jadi satu
df_index_member = pd.concat(df_list, ignore_index=True)

# Cek hasilnya
print(df_index_member.head())
print(df_index_member.columns)

folder_path = os.getcwd()  # Atau ubah ke folder lain
# Ambil semua file yang berakhiran .xlsx
excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
index = []
files = []
files_masuk = []
files_keluar = []
hasil_akhir = pd.DataFrame()

# Cetak daftar file .xlsx
for file in excel_files:
    if "Index Member" not in file:
        df = pd.read_excel(file)
        name = df.iloc[1, 6].replace(':',"").replace("-","").strip()
        evaluasi = df.iloc[2, 6].replace(':',"").strip()
        index.append(f"{name} {evaluasi}")
        
        table = []
        startrow = 0
        endrow = 0

        for i in range(len(df["Unnamed: 1"])):
            if df["Unnamed: 1"][i] == 1:
                table.append(startrow)
            try:
                if isinstance(df["Unnamed: 1"][i], int) and isinstance(df["Unnamed: 1"][i+1],str):
                    table.append(endrow+1)
            except:
                table.append(endrow+1)
            endrow+=1
            startrow+=1
        change= ["Jumlah Pra Evaluasi", "Jumlah Pasca Evaluasi", "Jumlah Keterangan", "Bobot Pra Evaluasi", "Bobot Pasca Evaluasi", "Bobot Keterangan"]
        header_masuk = [df.iloc[table[0]-2:table[0]-1, 1:4].values.tolist()[0] + change][0]
        masuk = df.iloc[table[0]:table[1], 1:12]
        masuk.columns = header_masuk
        exec(f'df_{name}_masuk = masuk')
        files.append(f'df_{name}_masuk')
        files_masuk.append(f'df_{name}_masuk')

        if len(table)>2:
            keluar = df.iloc[table[2]:table[3], 1:4]
            keluar.columns = df.iloc[table[2]-2:table[2]-1, 1:4].values.tolist()[0]
            exec(f'df_{name}_keluar = keluar')
            files.append(f'df_{name}_keluar')
            files_keluar.append(f'df_{name}_keluar')
    else:
        df_index_member = pd.read_excel(file)
        df_index_member["Index"] = df_index_member["Index"].str.replace("-", "", regex=True)

def Check_jumlah_naik(row):
    if row['Jumlah Keterangan'].upper() == "NAIK":
        return row['Jumlah Pasca Evaluasi'] > row['Jumlah Pra Evaluasi']
    if row['Jumlah Keterangan'].upper() == "TETAP":
        return round(row['Jumlah Pasca Evaluasi'],2) == round(row['Jumlah Pra Evaluasi'],2)
    if row['Jumlah Keterangan'].upper() == "TURUN":
        return row['Jumlah Pasca Evaluasi'] < row['Jumlah Pra Evaluasi']
    if row['Jumlah Keterangan'].upper() == "BARU":
        return (row['Jumlah Pasca Evaluasi'] > 0) & (str(row['Jumlah Pra Evaluasi']) == '-')
    return False

def Check_bobot_naik(row):
    if row['Bobot Keterangan'].upper() == "NAIK":
        return row['Bobot Pasca Evaluasi'] > row['Bobot Pra Evaluasi']
    if row['Bobot Keterangan'].upper() == "TETAP":
        return round(row['Bobot Pasca Evaluasi'],2) == round(row['Bobot Pra Evaluasi'],2)
    if row['Bobot Keterangan'].upper() == "TURUN":
        return row['Bobot Pasca Evaluasi'] < row['Bobot Pra Evaluasi']
    if row['Bobot Keterangan'].upper() == "BARU":
        return (row['Bobot Pasca Evaluasi'] > 0) & (str(row['Bobot Pra Evaluasi']) == '-')
    return False

for file in files_masuk:
    compare = file.split("_")[1]
    dict_compare = {'IDXECONOMIC30': 'ECONOMIC30'}
    if compare in dict_compare:
        compare = dict_compare[compare]
    current = df_index_member[df_index_member['Index'].str.upper() == compare]
    try:
        df_masuk = eval(file)
    except NameError:
        print(f"⚠️ DataFrame df_{file}_masuk tidak ditemukan.")
        continue 

    # Gabungkan berdasarkan kolom "Kode Efek" dan "Kode"
    exec(f'{file} = {file}.merge(current[["Kode Efek", "Weight Multiplier"]], left_on="Kode", right_on="Kode Efek", how="left")')
    exec(f'{file}["Check Kode"] = {file}["Kode Efek"] == {file}["Kode"]')
    exec(f'{file}["Check Jumlah"] = {file}["Weight Multiplier"] == {file}["Jumlah Pra Evaluasi"]')
    exec(f'{file}["Check Keterangan Jumlah"] = {file}.apply(Check_jumlah_naik, axis=1)')
    exec(f'{file}["Check Keterangan Bobot"] = {file}.apply(Check_bobot_naik, axis=1)')

    current = current.merge(df_masuk[['Kode','Jumlah Pra Evaluasi']], left_on="Kode Efek", right_on="Kode", how="left")
    current["Check Kode"] = current["Kode Efek"] == current["Kode"]
    current["Check Jumlah"] = current["Weight Multiplier"] == current["Jumlah Pra Evaluasi"]
    hasil_akhir = pd.concat([hasil_akhir, current], ignore_index=True)

    # Cek apakah DataFrame hasil kosong
    if current.empty:
        print(f"{file} not ok")

for file in files_keluar:
    compare = file.split("_")[1]
    dict_compare = {'IDXECONOMIC30': 'ECONOMIC30'}
    if compare in dict_compare:
        compare = dict_compare[compare]
    current = df_index_member[df_index_member['Index'].str.upper() == compare]
    try:
        df_masuk = eval(f"{file.replace("keluar","masuk")}")
    except NameError:
        print(f"⚠️ DataFrame df_{file}_keluar tidak ditemukan.")
        continue

    current = current.merge(df_masuk[['Kode']], left_on="Kode Efek", right_on="Kode", how="left")
    current["Check Kode"] = current["Kode Efek"] == current["Kode"]
    current = current[current['Check Kode'] == False]
    exec(f'{file} = {file}.merge(current[["Kode Efek"]], left_on="Kode", right_on="Kode Efek", how="left")')
    exec(f'{file}["Check Keluar"] = {file}["Kode"] == {file}["Kode Efek"]')

text = ""
text1 = ""
for naming in index:
    file = naming.split()[0]
    
    check_in = eval(f"df_{file}_masuk")
    # In
    in_baru = check_in[(check_in['Check Kode'] == False) & (check_in['Check Keterangan Jumlah'] == True) & (check_in['Check Keterangan Bobot'] == True)]
    if in_baru.shape[0] > 0:
        text += f"\n\n*{naming}:*"
        text += f"\n*- IN ({len(in_baru)}):*"
        for name in in_baru.Kode.unique().tolist():
            text += f' {name}'


    # In salah format
    in_baru_salah = check_in[((check_in['Check Kode'] == False) & (check_in['Check Keterangan Jumlah'] == False)) | ((check_in['Check Kode'] == False) & (check_in['Check Keterangan Bobot'] == False))]
    if in_baru_salah.shape[0] > 0:
        text += f"\n*- IN Error ({len(in_baru_salah)}):*"
        for name in in_baru_salah.Kode.unique().tolist():
            text += f' {name}'

    # Keluar
    if f"df_{file}_keluar" in files:
        check_out = eval(f"df_{file}_keluar")
        out_baru = check_out[check_out['Check Keluar'] == True]
        if out_baru.shape[0] > 0:
            text += f"\n*- OUT ({len(out_baru)}):*"
            for name in out_baru.Kode.unique().tolist():
                text += f' {name}'

    # Check Salah
    salah_in = check_in[((check_in['Check Kode'] == True) & (check_in['Check Keterangan Jumlah'] == False)) | ((check_in['Check Kode'] == True) & (check_in['Check Keterangan Bobot'] == False))]
    if salah_in.shape[0] > 0:
        text += f"\n*- Error ({len(salah_in)}):*"
        for name in salah_in.Kode.unique().tolist():
            text += f' {name}'

    # Check anggota
    check_name = []
    for char in file:
        if char.isdigit(): 
            check_name.append(char)
    if check_name: 
        check_name = int(''.join(check_name))
        
        if check_name < len(check_in):
            if len(check_in) != check_name:
                text += f'\n- {file} constituent berlebih {len(check_in)-check_name}'

for naming in index:
    file = naming.split()[0]
    if file not in text:
        text1 += f"*{naming}:* Tidak ada perubahan keluar masuk indeks\n"

text = text + "\n\n" + text1
text = text[2:]

with open("hasil.txt", "w", encoding="utf-8") as file:
    file.write(text)