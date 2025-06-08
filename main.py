import random
import string
import qrcode
import pandas as pd
import os
from datetime import datetime

# مسیر فایل اکسل
excel_file = "certificate_codes_with_qr.xlsx"

# پوشه QR
output_dir = "qr_codes"
os.makedirs(output_dir, exist_ok=True)

# لینک پایه برای QR
BASE_URL = "http://localhost:8080/portfolio/6269-2/"

# تابع ساخت کد تصادفی ترکیب عدد و حرف
def generate_code(length=8):
    return ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))  # k کوچیک

# اگر فایل اکسل وجود دارد، بخوان
if os.path.exists(excel_file):
    df_existing = pd.read_excel(excel_file)
else:
    df_existing = pd.DataFrame()

# تولید اطلاعات گواهی جدید
code = generate_code()
link = BASE_URL + code
qr = qrcode.make(link)
qr_path = os.path.join(output_dir, f"{code}.png")
qr.save(qr_path)

# داده جدید
new_row = {
    "Certificate Code": code,
    "QR Code Image": qr_path,
    "Certificate Link": link,
    "Course Name": "دوره‌ی دائمی",
    "Grade": "خیلی خوب",
    "Status": "موفقیت‌آمیز",
    "Generated At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
}

# افزودن به اکسل
df_new = pd.DataFrame([new_row])
df_final = pd.concat([df_existing, df_new], ignore_index=True)
df_final.to_excel(excel_file, index=False)

print("🎉 گواهی‌نامه جدید ساخته شد!")
print("📎 کد:", code)
print("🔗 لینک:", link)
print("📂 QR ذخیره شد در:", qr_path)
