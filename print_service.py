from flask import Flask, request
import serial
import json
from flask_cors import CORS
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import serial.tools.list_ports
import queue
import os
from werkzeug.serving import make_server

class PrintServiceGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Zebra Print Service")
        self.root.geometry("600x800")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # Flask app setup
        self.app = Flask(__name__)
        CORS(self.app)
        self.server = None
        self.server_thread = None
        self.is_running = False
        
        # Message queue for thread-safe logging
        self.log_queue = queue.Queue()
        
        # Variables
        self.com_port = tk.StringVar(value="COM1")
        self.baud_rate = tk.StringVar(value="9600")
        self.api_key = tk.StringVar(value="")
        
        # Config file
        self.config_file = "zebra_config.json"
        self.load_config()
        
        self.create_widgets()
        
        # Start log consumer
        self.root.after(100, self.process_logs)
        
    def create_widgets(self):
        # Port Settings Section
        settings_frame = ttk.LabelFrame(self.root, text="Port Settings", padding=10)
        settings_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(settings_frame, text="COM Port:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        com_ports = [""] + [port.device for port in serial.tools.list_ports.comports()]
        self.com_select = ttk.Combobox(settings_frame, textvariable=self.com_port, values=com_ports)
        self.com_select.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(settings_frame, text="Baud Rate:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        baud_select = ttk.Combobox(settings_frame, textvariable=self.baud_rate, 
                                 values=["9600", "19200", "38400", "57600", "115200"])
        baud_select.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        
        ttk.Label(settings_frame, text="API Key:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        api_key_entry = ttk.Entry(settings_frame, textvariable=self.api_key, width=30, show="*")
        api_key_entry.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        
        refresh_btn = ttk.Button(settings_frame, text="Refresh Ports", command=self.refresh_ports)
        refresh_btn.grid(row=0, column=2, padx=5, pady=5)
        
        # Service Status Section
        status_frame = ttk.LabelFrame(self.root, text="Service Status", padding=10)
        status_frame.pack(fill="x", padx=10, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="Service Stopped")
        self.status_label.pack(side="left", padx=5)
        
        self.port_status = ttk.Label(status_frame, text="Port: Not Connected")
        self.port_status.pack(side="right", padx=5)
        
        # Control Buttons
        btn_frame = ttk.Frame(self.root, padding=10)
        btn_frame.pack(fill="x", padx=10, pady=5)
        
        self.start_btn = ttk.Button(btn_frame, text="Start Service", command=self.start_service)
        self.start_btn.pack(side="left", padx=5)
        
        self.stop_btn = ttk.Button(btn_frame, text="Stop Service", command=self.stop_service, state="disabled")
        self.stop_btn.pack(side="left", padx=5)
        
        save_btn = ttk.Button(btn_frame, text="Save Config", command=self.save_config)
        save_btn.pack(side="right", padx=5)

        # Test Print Section
        test_frame = ttk.LabelFrame(self.root, text="Test Print", padding=10)
        test_frame.pack(fill="x", padx=10, pady=5)

        # Input fields
        self.test_inputs = {}
        fields = [
            ("Made in:", "made_in", "TH"),
            ("Model:", "t1", "SCG123"),
            ("Barcode:", "t11", "123456789"),
            ("Product Name:", "t2", "Gold Ring"),
            ("Weight:", "t3", "3.5g"),
            ("Diamond:", "t4", "0.5ct"),
            ("Ruby:", "t5", "0.3ct"),
            ("Gold:", "t6", "18K"),
            ("Size:", "size", "7"),
        ]

        for i, (label, key, default) in enumerate(fields):
            ttk.Label(test_frame, text=label).grid(row=i, column=0, padx=5, pady=2, sticky="e")
            self.test_inputs[key] = ttk.Entry(test_frame, width=40)
            self.test_inputs[key].insert(0, default)
            self.test_inputs[key].grid(row=i, column=1, padx=5, pady=2, sticky="ew")

        test_btn = ttk.Button(test_frame, text="Test Print", command=self.test_print)
        test_btn.grid(row=len(fields), column=0, columnspan=2, pady=10)
        
        # Log Section
        log_frame = ttk.LabelFrame(self.root, text="Log", padding=10)
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=10)
        self.log_text.pack(fill="both", expand=True)
        
        # Add clear log button
        clear_log_btn = ttk.Button(log_frame, text="Clear Log", command=self.clear_log)
        clear_log_btn.pack(side="right", padx=5, pady=5)
        
    def refresh_ports(self):
        """Refresh the list of available COM ports"""
        com_ports = [""] + [port.device for port in serial.tools.list_ports.comports()]
        self.com_select['values'] = com_ports
        self.log("COM ports refreshed")
        
    def clear_log(self):
        """Clear the log text area"""
        self.log_text.delete(1.0, tk.END)
        
    def safe_log(self, message):
        """Thread-safe logging method"""
        self.log_queue.put(message)
    
    def process_logs(self):
        """Process logs from the queue"""
        try:
            while not self.log_queue.empty():
                message = self.log_queue.get_nowait()
                self.log(message)
        except Exception as e:
            print(f"Error processing logs: {e}")
        finally:
            # Schedule to run again
            self.root.after(100, self.process_logs)
        
    def log(self, message):
        """Add message to log text area"""
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        
    def start_service(self):
        """Start the Flask server"""
        if self.is_running:
            return
            
        # Validate COM port
        if not self.com_port.get():
            messagebox.showerror("Error", "Please select a COM port")
            return
            
        # Validate connection to printer
        try:
            ser = serial.Serial(
                port=self.com_port.get(),
                baudrate=int(self.baud_rate.get()),
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                bytesize=serial.EIGHTBITS,
                timeout=1
            )
            ser.close()
        except Exception as e:
            messagebox.showerror("Error", f"Cannot connect to printer: {str(e)}")
            return
         
        # Setup routes
        self.setup_routes()
            
        # Start server
        self.is_running = True
        self.server = make_server('127.0.0.1', 5000, self.app)
        self.server_thread = threading.Thread(target=self.server.serve_forever)
        self.server_thread.daemon = True
        self.server_thread.start()
        
        self.status_label.config(text="Service Running")
        self.port_status.config(text=f"Port: {self.com_port.get()}")
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.log("Service Started")
            
    def stop_service(self):
        """Stop the Flask server"""
        if not self.is_running:
            return
            
        self.is_running = False
        if self.server:
            self.server.shutdown()
            self.server = None
            
        self.status_label.config(text="Service Stopped")
        self.port_status.config(text="Port: Not Connected")
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.log("Service Stopped")
    
    def setup_routes(self):
        """Set up Flask routes"""
        @self.app.route('/test-connection', methods=['GET'])
        def test_connection():
            """Test if the service is running and can connect to the printer"""
            try:
                if not self.is_running:
                    return json.dumps({
                        "status": "error",
                        "message": "Service is not running",
                        "serviceStatus": "stopped"
                    }), 503
                    
                if not self.com_port.get():
                    return json.dumps({
                        "status": "error", 
                        "message": "No COM port selected",
                        "serviceStatus": "running",
                        "printerStatus": "not_connected"
                    }), 200
                    
                # Test printer connection
                try:
                    ser = serial.Serial(
                        port=self.com_port.get(),
                        baudrate=int(self.baud_rate.get()),
                        parity=serial.PARITY_NONE,
                        stopbits=serial.STOPBITS_ONE,
                        bytesize=serial.EIGHTBITS,
                        timeout=1
                    )
                    ser.close()
                    
                    return json.dumps({
                        "status": "success",
                        "message": "Printer connected",
                        "serviceStatus": "running",
                        "printerStatus": "connected",
                        "port": self.com_port.get(),
                        "baudRate": self.baud_rate.get()
                    }), 200
                    
                except Exception as e:
                    return json.dumps({
                        "status": "error",
                        "message": f"Printer connection error: {str(e)}",
                        "serviceStatus": "running", 
                        "printerStatus": "error"
                    }), 200
                    
            except Exception as e:
                self.safe_log(f"Test connection error: {str(e)}")
                return json.dumps({
                    "status": "error", 
                    "message": f"Service error: {str(e)}"
                }), 500
        
        @self.app.route('/status', methods=['GET'])
        def get_status():
            return json.dumps({
                "status": "running" if self.is_running else "stopped",
                "port": self.com_port.get() if self.com_port.get() else "Not selected"
            }), 200
        
        @self.app.route('/print', methods=['POST'])
        def print_data():
            try:
                # Check API key if configured
                if self.api_key.get():
                    auth_header = request.headers.get('Authorization')
                    if not auth_header or not auth_header.startswith('Bearer ') or auth_header[7:] != self.api_key.get():
                        self.safe_log("Unauthorized access attempt")
                        return json.dumps({"status": "error", "message": "Unauthorized"}), 401
                
                # Get and validate data
                data = request.get_json()
                if not data or 'data' not in data:
                    self.safe_log("Invalid request: missing data")
                    return json.dumps({"status": "error", "message": "Missing data"}), 400
                
                zpl_data = data.get('data', '')
                
                # Basic ZPL validation (optional, can be removed if it causes issues with complex ZPL)
                if not zpl_data.startswith('^XA') or not zpl_data.endswith('^XZ'):
                    self.safe_log("Invalid ZPL format")
                    return json.dumps({"status": "error", "message": "Invalid ZPL format"}), 400
                
                # Attempt to print
                if self.print_to_zebra(zpl_data):
                    return json.dumps({"status": "success"}), 200
                else:
                    return json.dumps({"status": "error", "message": "Print failed"}), 500
                    
            except Exception as e:
                self.safe_log(f"API Error: {str(e)}")
                return json.dumps({"status": "error", "message": str(e)}), 500
    
    def print_to_zebra(self, data):
        """Send ZPL data to the Zebra printer"""
        try:
            if not self.com_port.get():
                self.log("Error: No COM port selected")
                return False
                
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
            self.log(f"Printed ZPL data (length: {len(data)})")
            return True
        except Exception as e:
            self.log(f"Print Error: {str(e)}")
            return False
    
    def test_print(self):
        """Generate and print a test label"""
        try:
            # Validate inputs
            for key, entry in self.test_inputs.items():
                if not entry.get().strip():
                    messagebox.showwarning("Warning", f"Field '{key}' is empty")
                    return
            
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
    
    def load_config(self):
        """Load configuration from file"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.com_port.set(config.get('com_port', ''))
                    self.baud_rate.set(config.get('baud_rate', '9600'))
                    self.api_key.set(config.get('api_key', ''))
                    print("Configuration loaded")
        except Exception as e:
            print(f"Error loading config: {e}")
    
    def save_config(self):
        """Save configuration to file"""
        try:
            config = {
                'com_port': self.com_port.get(),
                'baud_rate': self.baud_rate.get(),
                'api_key': self.api_key.get()
            }
            with open(self.config_file, 'w') as f:
                json.dump(config, f)
            self.log("Configuration saved")
        except Exception as e:
            self.log(f"Error saving config: {str(e)}")
    
    def on_closing(self):
        """Handle window closing event"""
        if self.is_running:
            if messagebox.askyesno("Quit", "Service is still running. Stop service and exit?"):
                self.stop_service()
                self.root.destroy()
        else:
            self.root.destroy()
    
    def run(self):
        """Run the main application loop"""
        self.root.mainloop()

if __name__ == '__main__':
    service = PrintServiceGUI()
    service.run()