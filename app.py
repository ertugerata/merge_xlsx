import openpyxl
import pandas as pd
import glob
import os

# Ayarlar
path = './'
files = glob.glob(os.path.join(path, "*.xlsx")) 
output_file = "birlesmis_temiz_liste.xlsx"

# Standart Excel sarı dolgu rengi kodları
YELLOW_COLORS = ['FFFFFF00', 'FFFF00'] 

all_data = []
header = None

for f in files:
    # Çıktı dosyasının kendisini işlememesi için kontrol
    if os.path.basename(f) == output_file:
        continue
        
    print(f"{f} işleniyor...")
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb.active 

    current_file_rows = list(ws.rows)
    if not current_file_rows:
        continue

    # Her zaman ilk 10 sütunu hedefle
    target_column_count = 10

    # Başlığı belirle (Sadece ilk dosyadan, ilk 10 sütun)
    if header is None:
        header = [cell.value for cell in current_file_rows[0][:target_column_count]]

    # 2. satırdan itibaren verileri kontrol et
    for row in ws.iter_rows(min_row=2):
        # Satırın ilk hücresinin rengine bak
        fill_color = row[0].fill.start_color.rgb
        
        # Eğer renk sarı değilse listeye ilk 10 sütunu ekle
        if fill_color not in YELLOW_COLORS:
            # Sütun sayısı ne olursa olsun sadece ilk 10 hücreyi al
            row_values = [cell.value for cell in row[:target_column_count]]
            all_data.append(row_values)

# DataFrame oluştur
combined_df = pd.DataFrame(all_data, columns=header)

# Kaydet
combined_df.to_excel(output_file, index=False)

print(f"\nİşlem tamamlandı! Toplam {len(all_data)} satır birleştirildi.")
print(f"Her dosyadan ilk {target_column_count} sütun alındı.")
print(f"Sonuç dosyası: {output_file}")