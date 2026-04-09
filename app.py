import openpyxl
import pandas as pd
import glob
import os

# Ayarlar
path = './'  # Dosyaların olduğu klasör
files = glob.glob(os.path.join(path, "*.xlsx")) 
output_file = "birlesmis_temiz_liste.xlsx"

# Standart Excel sarı dolgu rengi kodları
YELLOW_COLORS = ['FFFFFF00', 'FFFF00'] 

all_data = []
header = None

for f in files:
    print(f"{f} işleniyor...")
    wb = openpyxl.load_workbook(f, data_only=True)
    ws = wb.active 

    current_file_rows = list(ws.rows)
    if not current_file_rows:
        continue

    # Başlığı al (Sadece ilk dosyadan)
    if header is None:
        # None olmayan hücreleri filtreleyerek gerçek başlık sayısını bulabilirsiniz
        # veya manuel olarak header = header[:10] diyebilirsiniz.
        header = [cell.value for cell in current_file_rows[0] if cell.value is not None]
    
    # Başlık sayısını referans alarak verileri oku
    column_count = len(header)

    # 2. satırdan itibaren verileri kontrol et
    for row in ws.iter_rows(min_row=2):
        # Satırın ilk hücresinin dolgu rengini kontrol et 
        fill_color = row[0].fill.start_color.rgb
        
        # Eğer renk sarı değilse listeye ekle
        if fill_color not in YELLOW_COLORS:
            # SADECE başlık sayısı kadar hücreyi al (Hatanın çözümü burası)
            row_values = [cell.value for cell in row[:column_count]]
            all_data.append(row_values)

# DataFrame oluştur ve kaydet
combined_df = pd.DataFrame(all_data, columns=header)
combined_df.to_excel(output_file, index=False)

print(f"\nİşlem tamamlandı! Toplam {len(all_data)} satır birleştirildi.")
print(f"Sonuç dosyası: {output_file}")