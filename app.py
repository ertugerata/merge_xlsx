import openpyxl
import pandas as pd
import glob
import os

# Ayarlar
path = './'  # Dosyaların olduğu klasör
files = glob.glob(os.path.join(path, "*.xlsx")) # .xlsx dosyaları için
output_file = "birlesmis_temiz_liste.xlsx"

# Standart Excel sarı dolgu rengi kodu genellikle 'FFFFFF00' veya 'FFFF00'dır.
# Bazı durumlarda tema renkleri farklılık gösterebilir.
YELLOW_COLORS = ['FFFFFF00', 'FFFF00'] 

all_data = []
header = None

for f in files:
    print(f"{f} işleniyor...")
    # Dosyayı openpyxl ile aç (data_only=True formül yerine sonucu okur)
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb.active # İlk sekmeyi seçer

    # Başlığı belirle (Sadece ilk dosyadan veya her dosyanın ilk satırı aynı olduğu için)
    current_file_rows = list(ws.rows)
    if not current_file_rows:
        continue

    if header is None:
        header = [cell.value for cell in current_file_rows[0]]

    # 2. satırdan itibaren (başlık hariç) verileri kontrol et
    for row in ws.iter_rows(min_row=2):
        # Satırın ilk hücresinin dolgu rengini kontrol et 
        # (Eğer tüm satır boyalıysa ilk hücreye bakmak yeterlidir)
        fill_color = row[0].fill.start_color.rgb
        
        # Eğer renk sarı değilse listeye ekle
        if fill_color not in YELLOW_COLORS:
            row_values = [cell.value for cell in row]
            all_data.append(row_values)

# DataFrame oluştur ve kaydet
combined_df = pd.DataFrame(all_data, columns=header)
combined_df.to_excel(output_file, index=False)

print(f"\nİşlem tamamlandı! Toplam {len(all_data)} satır birleştirildi.")
print(f"Sonuç dosyası: {output_file}")