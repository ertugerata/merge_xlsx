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
target_column_count = 10  # Her dosyadan hedeflenen sütun sayısı

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

    # Başlığı belirle (Sadece ilk dosyadan, ilk 10 sütun)
    # NOT: İlk dosyanın satırı 10 hücreden KISA olabilir (ör. 7 hücre).
    # Bu durumda slicing [:10] sadece var olan 7 hücreyi döner ve
    # header 10 uzunluğunda garanti edilmez. Bu yüzden eksikse None ile
    # 10 elemana tamamlıyoruz.
    if header is None:
        header = [cell.value for cell in current_file_rows[0][:target_column_count]]
        if len(header) < target_column_count:
            header += [None] * (target_column_count - len(header))

    # 2. satırdan itibaren verileri kontrol et
    for row in ws.iter_rows(min_row=2):
        # Satırın ilk hücresinin rengine bak
        fill_color = row[0].fill.start_color.rgb
        
        # Eğer renk sarı değilse listeye ilk 10 sütunu ekle
        if fill_color not in YELLOW_COLORS:
            # Sütun sayısı ne olursa olsun sadece ilk 10 hücreyi al
            row_values = [cell.value for cell in row[:target_column_count]]
            # Bu dosyanın satırı 10 hücreden kısaysa (az sütunlu dosya),
            # aynı şekilde 10'a tamamla ki header ile boyutu eşleşsin
            if len(row_values) < target_column_count:
                row_values += [None] * (target_column_count - len(row_values))
            all_data.append(row_values)

# DataFrame oluştur
combined_df = pd.DataFrame(all_data, columns=header)

# Kaydet
combined_df.to_excel(output_file, index=False)

print(f"\nİşlem tamamlandı! Toplam {len(all_data)} satır birleştirildi.")
print(f"Her dosyadan ilk {target_column_count} sütun alındı.")
print(f"Sonuç dosyası: {output_file}")