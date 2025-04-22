import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import pywhatkit
import os

class WhatsAppAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📤 WhatsApp Automation Tool")
        self.root.geometry("700x600")
        self.root.configure(bg="#f0f4f7")

        self.excel_contacts = None
        self.excel_marks = None
        self.images_dict = {}

        self.style = ttk.Style()
        self.style.configure('TNotebook.Tab', padding=[10, 5], font=('Segoe UI', 10, 'bold'))

        self.tab_control = ttk.Notebook(root)
        self.tab1 = ttk.Frame(self.tab_control)
        self.tab2 = ttk.Frame(self.tab_control)

        self.tab_control.add(self.tab1, text='🖼️ Send Marksheets')
        self.tab_control.add(self.tab2, text='💬 Send Text Messages')
        self.tab_control.pack(expand=1, fill="both")

        self.setup_tab1()
        self.setup_tab2()
        self.setup_log_area()

    def setup_tab1(self):
        frame = self.tab1
        ttk.Button(frame, text="📄 Upload Excel File", command=self.load_excel_tab1).pack(pady=10)
        ttk.Button(frame, text="🖼️ Upload Marksheet Images", command=self.load_images).pack(pady=10)

        caption_frame = ttk.LabelFrame(frame, text="Caption")
        caption_frame.pack(pady=10, padx=20, fill="x")
        self.caption_entry = ttk.Entry(caption_frame)
        self.caption_entry.pack(padx=10, pady=5, fill="x")

        ttk.Button(frame, text="🚀 Send Marksheets", command=self.send_marksheets).pack(pady=20)

    def setup_tab2(self):
        frame = self.tab2
        ttk.Button(frame, text="📄 Upload Contact Excel", command=self.load_contacts_excel).pack(pady=10)
        ttk.Button(frame, text="📄 Upload Marks Excel", command=self.load_marks_excel).pack(pady=10)

        details_frame = ttk.LabelFrame(frame, text="Message Details")
        details_frame.pack(pady=10, padx=20, fill="x")

        self.practical_no = tk.IntVar()
        self.teacher_name = tk.StringVar()
        self.date = tk.StringVar()

        ttk.Label(details_frame, text="Practical No:").pack(anchor='w', padx=10, pady=2)
        ttk.Entry(details_frame, textvariable=self.practical_no).pack(fill="x", padx=10)

        ttk.Label(details_frame, text="Date (DD/MM/YYYY):").pack(anchor='w', padx=10, pady=2)
        ttk.Entry(details_frame, textvariable=self.date).pack(fill="x", padx=10)

        ttk.Label(details_frame, text="Teacher Name:").pack(anchor='w', padx=10, pady=2)
        ttk.Entry(details_frame, textvariable=self.teacher_name).pack(fill="x", padx=10)

        ttk.Button(frame, text="📩 Send Text Messages", command=self.send_text_messages).pack(pady=20)

    def setup_log_area(self):
        log_frame = ttk.LabelFrame(self.root, text="📜 Status Log")
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        self.status = tk.Text(log_frame, height=10, wrap="word")
        self.status.pack(fill="both", expand=True)

    def log(self, message):
        self.status.insert(tk.END, message + "\n")
        self.status.see(tk.END)
        self.root.update()

    def load_excel_tab1(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_contacts = path
            self.log(f"✅ Excel File Loaded: {os.path.basename(path)}")

    def load_images(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Image Files", "*.jpg *.png")])
        if file_paths:
            self.images_dict = {os.path.splitext(os.path.basename(p))[0].strip(): p for p in file_paths}
            self.log(f"✅ {len(file_paths)} image(s) loaded.")

    def send_marksheets(self):
        caption = self.caption_entry.get().strip()
        if not self.excel_contacts or not self.images_dict or not caption:
            self.log("⚠️ Please load Excel, select images, and enter a caption.")
            return

        try:
            data = pd.read_excel(self.excel_contacts)
        except Exception as e:
            self.log(f"❌ Failed to read Excel file: {e}")
            return

        for idx, row in data.iterrows():
            roll_no = str(row['Roll No.']).strip()
            phone = str(row['Phone']).strip()

            image_path = self.images_dict.get(roll_no)
            if not image_path:
                self.log(f"❌ No image for Roll No: {roll_no}")
                continue

            self.log(f"📤 Sending to {phone} (Roll No: {roll_no})...")
            try:
                pywhatkit.sendwhats_image(f"+91{phone}", image_path, caption=caption, tab_close=True, close_time=5)
                self.log(f"✅ Sent to {phone}")
            except Exception as e:
                self.log(f"❌ Error sending to {phone}: {e}")

        self.log("🎉 All marksheets processed.")

    def load_contacts_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_contacts = path
            self.log(f"✅ Contact Excel Loaded: {os.path.basename(path)}")

    def load_marks_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            self.excel_marks = path
            self.log(f"✅ Marks Excel Loaded: {os.path.basename(path)}")

    def send_text_messages(self):
        if not self.excel_contacts or not self.excel_marks:
            self.log("⚠️ Please load both contact and marks Excel files.")
            return

        try:
            contacts_df = pd.read_excel(self.excel_contacts)
            marks_df = pd.read_excel(self.excel_marks)
        except Exception as e:
            self.log(f"❌ Failed to read Excel files: {e}")
            return

        contacts_map = dict(zip(contacts_df['Roll No.'], contacts_df['Phone']))

        for idx, row in marks_df.iterrows():
            roll_no = row['Roll No.']
            phone = contacts_map.get(roll_no)

            if not phone:
                self.log(f"❌ No contact for Roll No: {roll_no}")
                continue

            try:
                msg = f"Hi, Roll no. {roll_no} your marks for practical {self.practical_no.get()} dated on {self.date.get()} are Performance: {row['P']}, Viva: {row['V']}, Attendance: {row['A']} total: {row['Total']}. For any queries contact {self.teacher_name.get()}."
                self.log(f"💬 Sending to {phone} (Roll No: {roll_no})...")
                pywhatkit.sendwhatmsg_instantly(f"+91{phone}", msg, tab_close=True, close_time=5)
                self.log(f"✅ Sent to {phone}")
            except Exception as e:
                self.log(f"❌ Error sending to {phone}: {e}")

        self.log("🎉 All messages sent.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppAutomationApp(root)
    root.mainloop()
