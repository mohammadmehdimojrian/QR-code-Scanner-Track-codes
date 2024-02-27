import cv2
from pyzbar.pyzbar import decode
import pandas as pd
from tkinter import Tk, Label, Button, filedialog, StringVar, Entry, Text, Scrollbar, VERTICAL, END
import threading
import queue
from datetime import datetime, timedelta
import winsound  # Import winsound for playing notification sounds


class QRCodeApp:
    def __init__(self, root):
        self.root = root
        self.root.title('QR Code Scanner and Excel Checker')

        self.excel_data = None
        self.qr_code_queue = queue.Queue()
        self.scan_log = {}  # Changed to a dictionary to track last scan timestamp

        # Sound file paths
        self.sound_found = r"mixkit-correct-answer-tone-2870.wav"
        self.sound_not_found = r"mixkit-game-show-wrong-answer-buzz-950.wav"

        self.setup_ui()

    def setup_ui(self):
        Label(self.root, text='Select an Excel file and start scanning QR codes with your webcam.').pack()
        Button(self.root, text='Select Excel File', command=self.select_excel).pack()

        self.manual_entry = StringVar()
        Entry(self.root, textvariable=self.manual_entry).pack()
        Button(self.root, text='Check QR Code Manually', command=self.check_manual_entry).pack()

        self.start_webcam_button = Button(self.root, text='Start Webcam', command=self.start_webcam_thread,
                                          state='disabled')
        self.start_webcam_button.pack()

        self.result_text = StringVar()
        Label(self.root, textvariable=self.result_text).pack()

        self.log_text = Text(self.root, height=10, width=50)
        self.log_text.pack(side="left", fill="y")
        scrollbar = Scrollbar(self.root, command=self.log_text.yview)
        scrollbar.pack(side="left", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)

    def select_excel(self):
        excel_path = filedialog.askopenfilename(title='Select Excel File', filetypes=[('Excel Files', '*.xls *.xlsx')])
        if excel_path:
            self.excel_data = pd.read_excel(excel_path)
            self.start_webcam_button.config(state='normal')

    def start_webcam_thread(self):
        threading.Thread(target=self.start_webcam, daemon=True).start()

    def start_webcam(self):
        cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        while True:
            ret, frame = cap.read()
            if not ret:
                continue
            decoded_objects = decode(frame)
            for obj in decoded_objects:
                data = obj.data.decode('utf-8')
                self.process_qr_code(data)
            cv2.imshow("Webcam QR Code Scanner", frame)
            if cv2.waitKey(1) == 27:  # Exit on ESC
                break
        cap.release()
        cv2.destroyAllWindows()

    def process_qr_code(self, data):
        now = datetime.now()
        data = int(data)
        if data in self.scan_log and now - self.scan_log[data] < timedelta(minutes=15):
            message = f"Duplicate QR code's value {data} (not added due to redundancy)."
        else:
            self.scan_log[data] = now  # Update last scan time
            if self.check_value_in_excel(data):
                message = f"QR code's value {data} was found in the Excel file."
                self.play_sound(self.sound_found)  # Play sound for found QR code
            else:
                message = f"QR code's value {data} was not found in the Excel file."
                self.play_sound(self.sound_not_found)  # Play sound for not found QR code
            self.update_log_display(data, now.strftime("%Y-%m-%d %H:%M:%S"))
        self.qr_code_queue.put(message)

    def check_value_in_excel(self, value):
        if self.excel_data is not None:
            return int(value) in self.excel_data[self.excel_data.columns[2]].values
        return False

    def check_manual_entry(self):
        data = self.manual_entry.get()
        self.process_qr_code(data)

    def update_ui(self):
        try:
            while True:
                message = self.qr_code_queue.get_nowait()
                self.result_text.set(message)
        except queue.Empty:
            pass
        self.root.after(100, self.update_ui)

    def update_log_display(self, data, timestamp):
        self.log_text.insert(END, f"{timestamp}: {data}\n")
        self.log_text.see(END)  # Scroll to the bottom

    def play_sound(self, sound_file):
        """Play a sound from a given file."""
        try:
            winsound.PlaySound(sound_file, winsound.SND_FILENAME)
        except Exception as e:
            print(f"Error playing sound: {e}")


if __name__ == '__main__':
    root = Tk()
    app = QRCodeApp(root)
    root.after(100, app.update_ui)
    root.mainloop()
