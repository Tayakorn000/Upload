import customtkinter as ctk
from tkinter import messagebox
import os
import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# --- การตั้งค่าระดับ Senior ---
SPREADSHEET_NAME = "VanSale"  # ชื่อไฟล์ Google Sheet
WORKSHEET_NAME = "สรุปการส่งของ" # ชื่อหน้า (Tab) ที่จัดการ
CREDS_FILE = "credentials.json" # ไฟล์กุญแจ Google API

# โฟลเดอร์ที่แอปจะสร้างให้บน Desktop
BASE_FOLDER = os.path.join(os.path.expanduser("~"), "Desktop", "ระบบบัญชี_VanSale")
REPORT_FOLDER = os.path.join(BASE_FOLDER, "รายงาน_Excel")
BACKUP_FOLDER = os.path.join(BASE_FOLDER, "Backup_ก่อนลบ")

for folder in [REPORT_FOLDER, BACKUP_FOLDER]:
    os.makedirs(folder, exist_ok=True)

class AccountingApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # ตั้งค่าหน้าจอ
        self.title("Van Sale Accounting Support")
        self.geometry("600x550")
        ctk.set_appearance_mode("light") # หรือ "dark"
        ctk.set_default_color_theme("blue")

        # ส่วนหัว (Header)
        self.header_label = ctk.CTkLabel(self, text="ระบบจัดการข้อมูลฝ่ายบัญชี", font=ctk.CTkFont(size=24, weight="bold"))
        self.header_label.pack(pady=(30, 10))
        
        self.sub_label = ctk.CTkLabel(self, text="บริษัท ตั้งเจริญมีนบุรี จำกัด", font=ctk.CTkFont(size=14))
        self.sub_label.pack(pady=(0, 30))

        # --- กลุ่มที่ 1: จัดการไฟล์ USB ---
        self.frame1 = ctk.CTkFrame(self)
        self.frame1.pack(pady=10, padx=40, fill="x")
        
        self.btn_usb = ctk.CTkButton(self.frame1, text="📁  1. เสียบสาย USB และย้ายไฟล์จากมือถือ", 
                                     command=self.handle_usb, height=50, font=ctk.CTkFont(size=15))
        self.btn_usb.pack(pady=20, padx=20, fill="x")

        # --- กลุ่มที่ 2: โหลดข้อมูล ---
        self.frame2 = ctk.CTkFrame(self)
        self.frame2.pack(pady=10, padx=40, fill="x")
        
        self.btn_download = ctk.CTkButton(self.frame2, text="📥  2. โหลดข้อมูลสรุปเป็น Excel (ระบุวันที่)", 
                                          command=self.download_data, height=50, fg_color="#2ECC71", hover_color="#27AE60",
                                          font=ctk.CTkFont(size=15))
        self.btn_download.pack(pady=20, padx=20, fill="x")

        # --- กลุ่มที่ 3: ล้างข้อมูล (ปุ่มสีแดงอันตราย) ---
        self.btn_clear = ctk.CTkButton(self, text="⚠️  3. ล้างข้อมูลหน้า 'สรุปการส่งของ' บน Cloud", 
                                       command=self.safe_clear, height=50, fg_color="#E74C3C", hover_color="#C0392B",
                                       font=ctk.CTkFont(size=15))
        self.btn_clear.pack(pady=20, padx=60, fill="x")

        # แถบสถานะ
        self.status_label = ctk.CTkLabel(self, text="สถานะ: พร้อมใช้งาน", text_color="gray")
        self.status_label.pack(side="bottom", pady=20)

    def get_sheet(self):
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, scope)
        client = gspread.authorize(creds)
        return client.open(SPREADSHEET_NAME).worksheet(WORKSHEET_NAME)

    def handle_usb(self):
        """เปิดหน้าต่างประกบกันเพื่อให้บัญชีลากไฟล์ได้ง่ายที่สุด"""
        try:
            os.startfile('shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}') # This PC
            os.startfile(BASE_FOLDER) # โฟลเดอร์งานบน Desktop
            messagebox.showinfo("Senior Support", "กรุณาลากไฟล์จากมือถือเข้าสู่โฟลเดอร์ที่เตรียมไว้ให้")
        except Exception as e:
            messagebox.showerror("Error", f"ไม่สามารถเปิดหน้าต่างได้: {e}")

    def download_data(self, is_backup=False):
        """โหลดข้อมูลมาเป็นไฟล์ Excel พร้อมวันที่และเวลา"""
        try:
            self.status_label.configure(text="กำลังดึงข้อมูลจาก Cloud...", text_color="#2980B9")
            self.update()
            
            sheet = self.get_sheet()
            data = sheet.get_all_records()
            
            if not data:
                if not is_backup: messagebox.showwarning("Warning", "ไม่พบข้อมูลบน Cloud")
                return None

            df = pd.DataFrame(data)
            now = datetime.datetime.now().strftime("%Y-%m-%d_%H%M")
            
            folder = BACKUP_FOLDER if is_backup else REPORT_FOLDER
            prefix = "BACKUP_" if is_backup else "รายงาน_"
            file_name = f"{prefix}{WORKSHEET_NAME}_{now}.xlsx"
            full_path = os.path.join(folder, file_name)
            
            df.to_excel(full_path, index=False)
            
            if not is_backup:
                self.status_label.configure(text=f"โหลดไฟล์ {file_name} สำเร็จ", text_color="green")
                os.startfile(REPORT_FOLDER)
                messagebox.showinfo("Success", f"ดาวน์โหลดข้อมูลเรียบร้อย!\nไฟล์อยู่ที่: {file_name}")
            return full_path

        except Exception as e:
            messagebox.showerror("Error", f"การเชื่อมต่อล้มเหลว: {e}")
            self.status_label.configure(text="เกิดข้อผิดพลาด", text_color="red")
            return None

    def safe_clear(self):
        """ล้างข้อมูลแบบ Senior: Backup ให้ก่อน ถ้าสำเร็จค่อยลบ"""
        if not messagebox.askyesno("ยืนยันการลบ", "คุณต้องการล้างข้อมูลตารางใช่หรือไม่?\n(ระบบจะบันทึกไฟล์สำรองให้พนักงานอัตโนมัติ)"):
            return

        # 1. ทำการ Backup ก่อน
        backup_path = self.download_data(is_backup=True)
        
        if backup_path:
            try:
                sheet = self.get_sheet()
                last_row = len(sheet.get_all_values())
                if last_row > 1:
                    sheet.delete_rows(2, last_row)
                    messagebox.showinfo("สำเร็จ", f"ล้างข้อมูลเรียบร้อยและสำรองข้อมูลไว้ที่โฟลเดอร์ Backup แล้ว")
                    self.status_label.configure(text="ล้างข้อมูลและ Backup สำเร็จ", text_color="green")
                else:
                    messagebox.showinfo("Info", "ไม่มีข้อมูลให้ลบ")
            except Exception as e:
                messagebox.showerror("Error", f"ลบข้อมูลไม่สำเร็จ: {e}")
        else:
            messagebox.showerror("Error", "ไม่สามารถลบได้ เนื่องจากไม่สามารถสำรองข้อมูลลงเครื่องได้")

if __name__ == "__main__":
    app = AccountingApp()
    app.mainloop()