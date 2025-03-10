from flask import Flask, request
import serial
import json
from flask_cors import CORS
import tkinter as tk
from tkinter import ttk, scrolledtext
import threading
import serial.tools.list_ports

class PrintServiceGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Zebra Print Service")
        self.root.geometry("600x800")
        
        # Flask app setup
        self.app = Flask(__name__)
        CORS(self.app)
        self.server_thread = None
        self.is_running = False
        
        # Variables
        self.com_port = tk.StringVar(value="COM1")
        self.baud_rate = tk.StringVar(value="9600")
        
        self.create_widgets()
        
    def create_widgets(self):
        # Port Settings Section
        settings_frame = ttk.LabelFrame(self.root, text="Port Settings", padding=10)
        settings_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(settings_frame, text="COM Port:").grid(row=0, column=0, padx=5, pady=5)
        com_ports = [port.device for port in serial.tools.list_ports.comports()]
        com_select = ttk.Combobox(settings_frame, textvariable=self.com_port, values=com_ports)
        com_select.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, pady=5)
        baud_select = ttk.Combobox(settings_frame, textvariable=self.baud_rate, 
                                 values=["9600", "19200", "38400", "57600", "115200"])
        baud_select.grid(row=1, column=1, padx=5, pady=5)
        
        # Service Status Section
        status_frame = ttk.LabelFrame(self.root, text="Service Status", padding=10)
        status_frame.pack(fill="x", padx=10, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Service Stopped")
        self.status_label.pack()
        
        # Control Buttons
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        self.start_btn = ttk.Button(btn_frame, text="Start Service", command=self.start_service)
        self.start_btn.pack(side="left", padx=5)
        
        self.stop_btn = ttk.Button(btn_frame, text="Stop Service", command=self.stop_service, state="disabled")
        self.stop_btn.pack(side="left", padx=5)

        # Test Print Section
        test_frame = ttk.LabelFrame(self.root, text="Test Print", padding=10)
        test_frame.pack(fill="x", padx=10, pady=5)

        # Input fields
        self.test_inputs = {}
        fields = [
            ("Made in:", "made_in"),
            ("Model:", "t1"),
            ("Barcode:", "t11"),
            ("Product Name:", "t2"),
            ("Weight:", "t3"),
            ("Diamond:", "t4"),
            ("Ruby:", "t5"),
            ("Gold:", "t6"),
            ("Size:", "size"),
        ]

        for i, (label, key) in enumerate(fields):
            ttk.Label(test_frame, text=label).grid(row=i, column=0, padx=5, pady=2, sticky="e")
            self.test_inputs[key] = ttk.Entry(test_frame, width=40)
            self.test_inputs[key].grid(row=i, column=1, padx=5, pady=2, sticky="ew")

        test_btn = ttk.Button(test_frame, text="Test Print", command=self.test_print)
        test_btn.grid(row=len(fields), column=0, columnspan=2, pady=10)
        
        # Log Section
        log_frame = ttk.LabelFrame(self.root, text="Log", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill="both", expand=True)
        
    def test_print(self):
        try:
            zpl = "^XA^LL200^MD25^LT40^XZ"
            zpl += "^XA"
            
            # Header
            zpl += f"^FO252,15^A0N,20,18^FD{self.test_inputs['t1'].get()}^FS"
            
            # Barcode
            zpl += f"^FO248,35^BY1,3.0:1,25^BCN,Y,N,N^FD{self.test_inputs['t11'].get()}^FS"
            
            # Product Name
            zpl += f"^FO248,65^A0N,20,18^FD{self.test_inputs['t2'].get()}^FS"
            
            # Specifications
            zpl += f"^FO250,090^A0N,14,16,B^FD{self.test_inputs['t3'].get()} {self.test_inputs['size'].get()}^FS"
            
            # Additional Details
            zpl += f"^FO426,015^A0N,14,16,B^FD{self.test_inputs['t4'].get()}^FS"
            zpl += f"^FO426,030^A0N,14,16,B^FD{self.test_inputs['t5'].get()}^FS"
            zpl += f"^FO426,045^A0N,14,16,B^FD{self.test_inputs['t6'].get()}^FS"
            
            # Made in
            zpl += f"^FO025,050^A0N,15,15,B^FD{self.test_inputs['made_in'].get()}^FS"
            
            zpl += "^XZ"
            
            if self.print_to_zebra(zpl):
                self.log("Test print successful")
            else:
                self.log("Test print failed")
                
        except Exception as e:
            self.log(f"Test print error: {str(e)}")

    def log(self, message):
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        
    def start_service(self):
        if not self.is_running:
            self.is_running = True
            self.server_thread = threading.Thread(target=self.run_flask)
            self.server_thread.daemon = True
            self.server_thread.start()
            
            self.status_label.config(text="Service Running")
            self.start_btn.config(state="disabled")
            self.stop_btn.config(state="normal")
            self.log("Service Started")
            
    def stop_service(self):
        if self.is_running:
            self.is_running = False
            self.status_label.config(text="Service Stopped")
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            self.log("Service Stopped")
    
    def print_to_zebra(self, data):
        try:
            ser = serial.Serial(
                port=self.com_port.get(),
                baudrate=int(self.baud_rate.get()),
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                bytesize=serial.EIGHTBITS,
                timeout=1
            )
            
            ser.write(data.encode())
            ser.close()
            self.log(f"Printed: {data}")
            return True
        except Exception as e:
            self.log(f"Print Error: {str(e)}")
            return False
    
    def run_flask(self):
        @self.app.route('/print', methods=['POST'])
        def print_data():
            try:
                data = request.get_json()
                zpl_data = data.get('data', '')
                
                if self.print_to_zebra(zpl_data):
                    return json.dumps({"status": "success"}), 200
                else:
                    return json.dumps({"status": "error", "message": "Print failed"}), 500
                    
            except Exception as e:
                return json.dumps({"status": "error", "message": str(e)}), 500
        
        self.app.run(port=5000)
    
    def run(self):
        self.root.mainloop()

if __name__ == '__main__':
    service = PrintServiceGUI()
    service.run()