import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter import scrolledtext, filedialog, messagebox
import datetime
import re

try:
    import enchant
except ImportError:
    enchant = None

def add_context_menu(widget):
    """Adds a right-click context menu with cut, copy, and paste commands to the given widget."""
    menu = tk.Menu(widget, tearoff=0)
    menu.add_command(label="Cut", command=lambda: widget.event_generate("<<Cut>>"))
    menu.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"))
    menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
    
    def show_menu(event):
        menu.tk_popup(event.x_root, event.y_root)
    
    widget.bind("<Button-3>", show_menu)

class ReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Report Generator")
        
        # Initialize dictionary if pyenchant is available
        if enchant:
            self.dictionary = enchant.Dict("en_US")
        else:
            self.dictionary = None
        
        # Report Number (Row 0)
        tk.Label(root, text="Report Number:").grid(row=0, column=0, sticky='w')
        self.report_number = ttk.Entry(root)
        self.report_number.grid(row=0, column=1, padx=10, pady=5)
        add_context_menu(self.report_number)
        
        # Executive Summary (Row 1)
        tk.Label(root, text="Executive Summary:").grid(row=1, column=0, sticky='w')
        self.summary_text = scrolledtext.ScrolledText(root, width=60, height=10)
        self.summary_text.grid(row=1, column=1, padx=10, pady=5)
        add_context_menu(self.summary_text)
        self.summary_text.tag_config("misspelled", foreground="red", underline=True, background="yellow")
        self.summary_text.tag_bind("misspelled", "<Button-3>", self.on_misspelled_right_click)
        
        # Spell Check Button (Row 1, Column 2)
        self.spell_check_button = ttk.Button(root, text="Spell Check", command=self.spell_check_executive_summary)
        self.spell_check_button.grid(row=1, column=2, padx=5, pady=5, sticky='n')
        
        # Item Request Date (Row 2)
        tk.Label(root, text="Item Request Date:").grid(row=2, column=0, sticky='w')
        self.drop_off_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.drop_off_date.grid(row=2, column=1, padx=10, pady=5)
        
        # Dropoff Person (Row 3)
        tk.Label(root, text="Dropoff Person:").grid(row=3, column=0, sticky='w')
        self.drop_off_person = ttk.Entry(root)
        self.drop_off_person.grid(row=3, column=1, padx=10, pady=5)
        add_context_menu(self.drop_off_person)
        
        # Dropoff Date (Row 4)
        tk.Label(root, text="Dropoff Date:").grid(row=4, column=0, sticky='w')
        self.dropoff_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.dropoff_date.grid(row=4, column=1, padx=10, pady=5)
        
        # Legal Basis (Row 5)
        tk.Label(root, text="Legal Basis:").grid(row=5, column=0, sticky='w')
        self.legal_basis = ttk.Combobox(root, values=["Search Warrant", "Consent", "Abandoned Property", "Deceased Person"])
        self.legal_basis.grid(row=5, column=1, padx=10, pady=5)
        
        # Program Used (Row 6) - Using Checkbuttons for multiple selections
        tk.Label(root, text="Program Used:").grid(row=6, column=0, sticky='w')
        self.program_used_vars = {}
        program_options = ["Cellebrite", "GrayKey", "FTK Imager", "FEX", "FTK", "Other"]
        program_frame = tk.Frame(root)
        program_frame.grid(row=6, column=1, padx=10, pady=5, sticky='w')
        for option in program_options:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(program_frame, text=option, variable=var)
            cb.pack(side=tk.LEFT, padx=5)
            self.program_used_vars[option] = var
        
        # Decoder Used (Row 7) - Using Checkbuttons for multiple selections
        tk.Label(root, text="Decoder Used:").grid(row=7, column=0, sticky='w')
        self.decoder_used_vars = {}
        decoder_options = ["Physical Analyzer", "Axiom", "FEX", "X-Ways", "FTK", "Other"]
        decoder_frame = tk.Frame(root)
        decoder_frame.grid(row=7, column=1, padx=10, pady=5, sticky='w')
        for option in decoder_options:
            var = tk.BooleanVar()
            cb = tk.Checkbutton(decoder_frame, text=option, variable=var)
            cb.pack(side=tk.LEFT, padx=5)
            self.decoder_used_vars[option] = var
        
        # MD5 Hash Values (Row 8)
        tk.Label(root, text="MD5 Hash Values:").grid(row=8, column=0, sticky='w')
        self.md5_hash = scrolledtext.ScrolledText(root, width=60, height=5)
        self.md5_hash.grid(row=8, column=1, padx=10, pady=5)
        add_context_menu(self.md5_hash)
        
        # Disposition Date (Row 9)
        tk.Label(root, text="Disposition Date:").grid(row=9, column=0, sticky='w')
        self.disposition_date = DateEntry(root, width=12, background='darkblue', foreground='white', borderwidth=2)
        self.disposition_date.grid(row=9, column=1, padx=10, pady=5)
        
        # Disposition (Row 10)
        tk.Label(root, text="Disposition:").grid(row=10, column=0, sticky='w')
        self.disposition = ttk.Combobox(root, values=[
            "returned to the lead Investigator.",
            "returned to the owner.",
            "placed into evidence",
            "retained for brute forcing or future support."
        ])
        self.disposition.grid(row=10, column=1, padx=10, pady=5)
        
        # Devices Section: Add Device Button (Row 11)
        self.devices = []
        self.add_device_button = ttk.Button(root, text="Add Device", command=self.add_device)
        self.add_device_button.grid(row=11, column=1, pady=10)
        
        # Display Saved Devices (Row 12)
        tk.Label(root, text="Saved Devices:").grid(row=12, column=0, sticky='w')
        self.devices_listbox = tk.Listbox(root, width=100, height=10)
        self.devices_listbox.grid(row=12, column=1, padx=10, pady=5)
        self.devices_listbox.bind('<Double-Button-1>', self.edit_device)
        
        # Generate Report Button (Row 13)
        self.generate_button = ttk.Button(root, text="Generate Report", command=self.generate_report)
        self.generate_button.grid(row=13, column=1, pady=10)

    def spell_check_executive_summary(self):
        if self.dictionary is None:
            messagebox.showerror("Spell Check Error", "pyenchant is not installed.\nInstall it using 'pip install pyenchant'")
            return

        text_widget = self.summary_text
        content = text_widget.get("1.0", tk.END)
        text_widget.tag_remove("misspelled", "1.0", tk.END)
        misspelled_count = 0
        for match in re.finditer(r"\b\w+\b", content):
            word = match.group()
            if not self.dictionary.check(word):
                start_index = f"1.0+{match.start()}c"
                end_index = f"1.0+{match.end()}c"
                text_widget.tag_add("misspelled", start_index, end_index)
                misspelled_count += 1
        messagebox.showinfo("Spell Check", f"Spell check complete.\nMisspelled words found: {misspelled_count}")

    def on_misspelled_right_click(self, event):
        if self.dictionary is None:
            return
        widget = event.widget
        index = widget.index(f"@{event.x},{event.y}")
        word_start = widget.index(f"{index} wordstart")
        word_end = widget.index(f"{index} wordend")
        word = widget.get(word_start, word_end)
        suggestions = self.dictionary.suggest(word)
        menu = tk.Menu(widget, tearoff=0)
        if suggestions:
            for suggestion in suggestions[:5]:
                menu.add_command(label=suggestion, command=lambda s=suggestion: self.replace_word(widget, word_start, word_end, s))
        else:
            menu.add_command(label="No suggestions", state=tk.DISABLED)
        menu.tk_popup(event.x_root, event.y_root)

    def replace_word(self, widget, start, end, replacement):
        widget.delete(start, end)
        widget.insert(start, replacement)
        widget.tag_remove("misspelled", start, f"{start}+{len(replacement)}c")

    def add_device(self, device_info=None, index=None):
        device_window = tk.Toplevel(self.root)
        device_window.title("Add Device" if device_info is None else "Edit Device")
        
        # Common field: Guardian Item Number (for all devices)
        tk.Label(device_window, text="Guardian Item Number:").grid(row=0, column=0, sticky="w")
        guardian_entry = ttk.Entry(device_window)
        guardian_entry.grid(row=0, column=1, padx=10, pady=5)
        add_context_menu(guardian_entry)
        
        # Device Type Selection
        tk.Label(device_window, text="Device Type:").grid(row=1, column=0, sticky="w")
        device_types = ["Mobile Device", "Hard Drive", "SSD", "Drone", "Infotainment System", "Desktop Computer", "Laptop", "All Other"]
        device_type_cb = ttk.Combobox(device_window, values=device_types, state="readonly")
        device_type_cb.grid(row=1, column=1, padx=10, pady=5)
        device_type_cb.set("Mobile Device")
        
        # Container for device-type–specific fields
        fields_container = tk.Frame(device_window)
        fields_container.grid(row=2, column=0, columnspan=2, padx=10, pady=5)
        frames = {}
        
        # ---------------- Mobile Device Frame ----------------
        mobile_frame = tk.Frame(fields_container)
        frames["Mobile Device"] = mobile_frame
        tk.Label(mobile_frame, text="BFU/AFU:").grid(row=0, column=0, sticky="w")
        mobile_state = ttk.Combobox(mobile_frame, values=["HOT", "COLD", "UNKNOWN"])
        mobile_state.grid(row=0, column=1, padx=5, pady=2)
        tk.Label(mobile_frame, text="Airplane Mode:").grid(row=1, column=0, sticky="w")
        mobile_airplane = ttk.Combobox(mobile_frame, values=["Yes", "No", "Unable to Determine"])
        mobile_airplane.grid(row=1, column=1, padx=5, pady=2)
        tk.Label(mobile_frame, text="Micro SD Present:").grid(row=2, column=0, sticky="w")
        mobile_sd = ttk.Combobox(mobile_frame, values=["Yes", "No", "Unable to Determine"])
        mobile_sd.grid(row=2, column=1, padx=5, pady=2)
        tk.Label(mobile_frame, text="Make:").grid(row=3, column=0, sticky="w")
        mobile_make = ttk.Combobox(mobile_frame, values=["Apple", "Samsung", "Huawei", "LG", "Lenovo", "Motorola", "Xiaomi", "Oppo", "Vivo", "Google", "Sony Eperia", "OnePlus", "Nokia", "Asus", "HTC", "ZTE", "Meizu", "Blackberry", "Other"])
        mobile_make.grid(row=3, column=1, padx=5, pady=2)
        tk.Label(mobile_frame, text="Model:").grid(row=4, column=0, sticky="w")
        mobile_model = ttk.Entry(mobile_frame)
        mobile_model.grid(row=4, column=1, padx=5, pady=2)
        add_context_menu(mobile_model)
        tk.Label(mobile_frame, text="SIM Present:").grid(row=5, column=0, sticky="w")
        mobile_sim = ttk.Combobox(mobile_frame, values=["Yes", "No", "Unknown"])
        mobile_sim.grid(row=5, column=1, padx=5, pady=2)
        tk.Label(mobile_frame, text="ICCID:").grid(row=6, column=0, sticky="w")
        mobile_iccid = ttk.Entry(mobile_frame)
        mobile_iccid.grid(row=6, column=1, padx=5, pady=2)
        add_context_menu(mobile_iccid)
        tk.Label(mobile_frame, text="Phone Number:").grid(row=7, column=0, sticky="w")
        mobile_phone = ttk.Entry(mobile_frame)
        mobile_phone.grid(row=7, column=1, padx=5, pady=2)
        add_context_menu(mobile_phone)
        tk.Label(mobile_frame, text="Serial Number:").grid(row=8, column=0, sticky="w")
        mobile_serial = ttk.Entry(mobile_frame)
        mobile_serial.grid(row=8, column=1, padx=5, pady=2)
        add_context_menu(mobile_serial)
        tk.Label(mobile_frame, text="IMEI:").grid(row=9, column=0, sticky="w")
        mobile_imei = ttk.Entry(mobile_frame)
        mobile_imei.grid(row=9, column=1, padx=5, pady=2)
        add_context_menu(mobile_imei)
        tk.Label(mobile_frame, text="Acquisition Type:").grid(row=10, column=0, sticky="w")
        mobile_acq = ttk.Combobox(mobile_frame, values=["Fully Physical", "Full File System", "AFU File System", "BFU File System", "Advanced Logical", "Logical", "Camera"])
        mobile_acq.grid(row=10, column=1, padx=5, pady=2)
        
        # ---------------- Hard Drive Frame ----------------
        hd_frame = tk.Frame(fields_container)
        frames["Hard Drive"] = hd_frame
        tk.Label(hd_frame, text="Brand:").grid(row=0, column=0, sticky="w")
        hd_brand = ttk.Entry(hd_frame)
        hd_brand.grid(row=0, column=1, padx=5, pady=2)
        add_context_menu(hd_brand)
        tk.Label(hd_frame, text="Capacity:").grid(row=1, column=0, sticky="w")
        hd_capacity = ttk.Entry(hd_frame)
        hd_capacity.grid(row=1, column=1, padx=5, pady=2)
        add_context_menu(hd_capacity)
        tk.Label(hd_frame, text="Interface Type:").grid(row=2, column=0, sticky="w")
        hd_interface = ttk.Combobox(hd_frame, values=["SATA", "NVMe", "IDE", "Other"])
        hd_interface.grid(row=2, column=1, padx=5, pady=2)
        tk.Label(hd_frame, text="Model:").grid(row=3, column=0, sticky="w")
        hd_model = ttk.Entry(hd_frame)
        hd_model.grid(row=3, column=1, padx=5, pady=2)
        add_context_menu(hd_model)
        tk.Label(hd_frame, text="Serial Number:").grid(row=4, column=0, sticky="w")
        hd_serial = ttk.Entry(hd_frame)
        hd_serial.grid(row=4, column=1, padx=5, pady=2)
        add_context_menu(hd_serial)
        tk.Label(hd_frame, text="Firmware Version:").grid(row=5, column=0, sticky="w")
        hd_firmware = ttk.Entry(hd_frame)
        hd_firmware.grid(row=5, column=1, padx=5, pady=2)
        add_context_menu(hd_firmware)
        tk.Label(hd_frame, text="Condition:").grid(row=6, column=0, sticky="w")
        hd_condition = ttk.Combobox(hd_frame, values=["Working", "Damaged", "Unknown"])
        hd_condition.grid(row=6, column=1, padx=5, pady=2)
        
        # ---------------- SSD Frame ----------------
        ssd_frame = tk.Frame(fields_container)
        frames["SSD"] = ssd_frame
        tk.Label(ssd_frame, text="Brand:").grid(row=0, column=0, sticky="w")
        ssd_brand = ttk.Entry(ssd_frame)
        ssd_brand.grid(row=0, column=1, padx=5, pady=2)
        add_context_menu(ssd_brand)
        tk.Label(ssd_frame, text="Capacity:").grid(row=1, column=0, sticky="w")
        ssd_capacity = ttk.Entry(ssd_frame)
        ssd_capacity.grid(row=1, column=1, padx=5, pady=2)
        add_context_menu(ssd_capacity)
        tk.Label(ssd_frame, text="Interface Type:").grid(row=2, column=0, sticky="w")
        ssd_interface = ttk.Combobox(ssd_frame, values=["SATA", "NVMe", "M.2", "Other"])
        ssd_interface.grid(row=2, column=1, padx=5, pady=2)
        tk.Label(ssd_frame, text="Model:").grid(row=3, column=0, sticky="w")
        ssd_model = ttk.Entry(ssd_frame)
        ssd_model.grid(row=3, column=1, padx=5, pady=2)
        add_context_menu(ssd_model)
        tk.Label(ssd_frame, text="Serial Number:").grid(row=4, column=0, sticky="w")
        ssd_serial = ttk.Entry(ssd_frame)
        ssd_serial.grid(row=4, column=1, padx=5, pady=2)
        add_context_menu(ssd_serial)
        tk.Label(ssd_frame, text="Firmware Version:").grid(row=5, column=0, sticky="w")
        ssd_firmware = ttk.Entry(ssd_frame)
        ssd_firmware.grid(row=5, column=1, padx=5, pady=2)
        add_context_menu(ssd_firmware)
        tk.Label(ssd_frame, text="Condition:").grid(row=6, column=0, sticky="w")
        ssd_condition = ttk.Combobox(ssd_frame, values=["Working", "Damaged", "Unknown"])
        ssd_condition.grid(row=6, column=1, padx=5, pady=2)
        
        # ---------------- Drone Frame ----------------
        drone_frame = tk.Frame(fields_container)
        frames["Drone"] = drone_frame
        tk.Label(drone_frame, text="Manufacturer:").grid(row=0, column=0, sticky="w")
        drone_manufacturer = ttk.Combobox(drone_frame, values=["DJI", "Parrot", "Autel", "Other"])
        drone_manufacturer.grid(row=0, column=1, padx=5, pady=2)
        tk.Label(drone_frame, text="Model:").grid(row=1, column=0, sticky="w")
        drone_model = ttk.Entry(drone_frame)
        drone_model.grid(row=1, column=1, padx=5, pady=2)
        add_context_menu(drone_model)
        tk.Label(drone_frame, text="Serial Number:").grid(row=2, column=0, sticky="w")
        drone_serial = ttk.Entry(drone_frame)
        drone_serial.grid(row=2, column=1, padx=5, pady=2)
        add_context_menu(drone_serial)
        tk.Label(drone_frame, text="Firmware Version:").grid(row=3, column=0, sticky="w")
        drone_firmware = ttk.Entry(drone_frame)
        drone_firmware.grid(row=3, column=1, padx=5, pady=2)
        add_context_menu(drone_firmware)
        tk.Label(drone_frame, text="Battery Health:").grid(row=4, column=0, sticky="w")
        drone_battery = ttk.Entry(drone_frame)
        drone_battery.grid(row=4, column=1, padx=5, pady=2)
        add_context_menu(drone_battery)
        tk.Label(drone_frame, text="Camera Resolution:").grid(row=5, column=0, sticky="w")
        drone_camera = ttk.Entry(drone_frame)
        drone_camera.grid(row=5, column=1, padx=5, pady=2)
        add_context_menu(drone_camera)
        tk.Label(drone_frame, text="Acquisition Type:").grid(row=6, column=0, sticky="w")
        drone_acq = ttk.Entry(drone_frame)
        drone_acq.grid(row=6, column=1, padx=5, pady=2)
        add_context_menu(drone_acq)
        
        # ---------------- Infotainment System Frame ----------------
        infotainment_frame = tk.Frame(fields_container)
        frames["Infotainment System"] = infotainment_frame
        tk.Label(infotainment_frame, text="Manufacturer:").grid(row=0, column=0, sticky="w")
        info_manuf = ttk.Entry(infotainment_frame)
        info_manuf.grid(row=0, column=1, padx=5, pady=2)
        add_context_menu(info_manuf)
        tk.Label(infotainment_frame, text="Model:").grid(row=1, column=0, sticky="w")
        info_model = ttk.Entry(infotainment_frame)
        info_model.grid(row=1, column=1, padx=5, pady=2)
        add_context_menu(info_model)
        tk.Label(infotainment_frame, text="Software Version:").grid(row=2, column=0, sticky="w")
        info_soft = ttk.Entry(infotainment_frame)
        info_soft.grid(row=2, column=1, padx=5, pady=2)
        add_context_menu(info_soft)
        tk.Label(infotainment_frame, text="Serial Number:").grid(row=3, column=0, sticky="w")
        info_serial = ttk.Entry(infotainment_frame)
        info_serial.grid(row=3, column=1, padx=5, pady=2)
        add_context_menu(info_serial)
        tk.Label(infotainment_frame, text="Firmware Version:").grid(row=4, column=0, sticky="w")
        info_firmware = ttk.Entry(infotainment_frame)
        info_firmware.grid(row=4, column=1, padx=5, pady=2)
        add_context_menu(info_firmware)
        tk.Label(infotainment_frame, text="Vehicle Make:").grid(row=5, column=0, sticky="w")
        info_vehicle_make = ttk.Entry(infotainment_frame)
        info_vehicle_make.grid(row=5, column=1, padx=5, pady=2)
        add_context_menu(info_vehicle_make)
        tk.Label(infotainment_frame, text="Vehicle Model:").grid(row=6, column=0, sticky="w")
        info_vehicle_model = ttk.Entry(infotainment_frame)
        info_vehicle_model.grid(row=6, column=1, padx=5, pady=2)
        add_context_menu(info_vehicle_model)
        
        # ---------------- Desktop Computer Frame ----------------
        desktop_frame = tk.Frame(fields_container)
        frames["Desktop Computer"] = desktop_frame
        tk.Label(desktop_frame, text="Manufacturer:").grid(row=0, column=0, sticky="w")
        desk_manuf = ttk.Entry(desktop_frame)
        desk_manuf.grid(row=0, column=1, padx=5, pady=2)
        add_context_menu(desk_manuf)
        tk.Label(desktop_frame, text="Model:").grid(row=1, column=0, sticky="w")
        desk_model = ttk.Entry(desktop_frame)
        desk_model.grid(row=1, column=1, padx=5, pady=2)
        add_context_menu(desk_model)
        tk.Label(desktop_frame, text="Serial Number:").grid(row=2, column=0, sticky="w")
        desk_serial = ttk.Entry(desktop_frame)
        desk_serial.grid(row=2, column=1, padx=5, pady=2)
        add_context_menu(desk_serial)
        tk.Label(desktop_frame, text="Operating System:").grid(row=3, column=0, sticky="w")
        desk_os = ttk.Entry(desktop_frame)
        desk_os.grid(row=3, column=1, padx=5, pady=2)
        add_context_menu(desk_os)
        tk.Label(desktop_frame, text="CPU:").grid(row=4, column=0, sticky="w")
        desk_cpu = ttk.Entry(desktop_frame)
        desk_cpu.grid(row=4, column=1, padx=5, pady=2)
        add_context_menu(desk_cpu)
        tk.Label(desktop_frame, text="RAM:").grid(row=5, column=0, sticky="w")
        desk_ram = ttk.Entry(desktop_frame)
        desk_ram.grid(row=5, column=1, padx=5, pady=2)
        add_context_menu(desk_ram)
        tk.Label(desktop_frame, text="Storage Capacity:").grid(row=6, column=0, sticky="w")
        desk_storage = ttk.Entry(desktop_frame)
        desk_storage.grid(row=6, column=1, padx=5, pady=2)
        add_context_menu(desk_storage)
        tk.Label(desktop_frame, text="Acquisition Type:").grid(row=7, column=0, sticky="w")
        desk_acq = ttk.Entry(desktop_frame)
        desk_acq.grid(row=7, column=1, padx=5, pady=2)
        add_context_menu(desk_acq)
        
        # ---------------- Laptop Frame ----------------
        laptop_frame = tk.Frame(fields_container)
        frames["Laptop"] = laptop_frame
        tk.Label(laptop_frame, text="Manufacturer:").grid(row=0, column=0, sticky="w")
        lap_manuf = ttk.Entry(laptop_frame)
        lap_manuf.grid(row=0, column=1, padx=5, pady=2)
        add_context_menu(lap_manuf)
        tk.Label(laptop_frame, text="Model:").grid(row=1, column=0, sticky="w")
        lap_model = ttk.Entry(laptop_frame)
        lap_model.grid(row=1, column=1, padx=5, pady=2)
        add_context_menu(lap_model)
        tk.Label(laptop_frame, text="Serial Number:").grid(row=2, column=0, sticky="w")
        lap_serial = ttk.Entry(laptop_frame)
        lap_serial.grid(row=2, column=1, padx=5, pady=2)
        add_context_menu(lap_serial)
        tk.Label(laptop_frame, text="Operating System:").grid(row=3, column=0, sticky="w")
        lap_os = ttk.Entry(laptop_frame)
        lap_os.grid(row=3, column=1, padx=5, pady=2)
        add_context_menu(lap_os)
        tk.Label(laptop_frame, text="CPU:").grid(row=4, column=0, sticky="w")
        lap_cpu = ttk.Entry(laptop_frame)
        lap_cpu.grid(row=4, column=1, padx=5, pady=2)
        add_context_menu(lap_cpu)
        tk.Label(laptop_frame, text="RAM:").grid(row=5, column=0, sticky="w")
        lap_ram = ttk.Entry(laptop_frame)
        lap_ram.grid(row=5, column=1, padx=5, pady=2)
        add_context_menu(lap_ram)
        tk.Label(laptop_frame, text="Storage Capacity:").grid(row=6, column=0, sticky="w")
        lap_storage = ttk.Entry(laptop_frame)
        lap_storage.grid(row=6, column=1, padx=5, pady=2)
        add_context_menu(lap_storage)
        tk.Label(laptop_frame, text="Battery Health:").grid(row=7, column=0, sticky="w")
        lap_battery = ttk.Entry(laptop_frame)
        lap_battery.grid(row=7, column=1, padx=5, pady=2)
        add_context_menu(lap_battery)
        
        # ---------------- All Other Frame ----------------
        other_frame = tk.Frame(fields_container)
        frames["All Other"] = other_frame
        tk.Label(other_frame, text="Device Description:").grid(row=0, column=0, sticky="w")
        other_desc = scrolledtext.ScrolledText(other_frame, width=40, height=4)
        other_desc.grid(row=0, column=1, padx=5, pady=2)
        add_context_menu(other_desc)
        tk.Label(other_frame, text="Make:").grid(row=1, column=0, sticky="w")
        other_make = ttk.Entry(other_frame)
        other_make.grid(row=1, column=1, padx=5, pady=2)
        add_context_menu(other_make)
        tk.Label(other_frame, text="Model:").grid(row=2, column=0, sticky="w")
        other_model = ttk.Entry(other_frame)
        other_model.grid(row=2, column=1, padx=5, pady=2)
        add_context_menu(other_model)
        tk.Label(other_frame, text="Serial Number:").grid(row=3, column=0, sticky="w")
        other_serial = ttk.Entry(other_frame)
        other_serial.grid(row=3, column=1, padx=5, pady=2)
        add_context_menu(other_serial)
        
        def show_frame(selected):
            for f in frames.values():
                f.grid_forget()
            frames[selected].grid(row=0, column=0, sticky="w")
        
        device_type_cb.bind("<<ComboboxSelected>>", lambda e: show_frame(device_type_cb.get()))
        show_frame(device_type_cb.get())
        
        def save_device():
            selected_type = device_type_cb.get()
            data = {"Device Type": selected_type, "Guardian Item Number": guardian_entry.get()}
            if not data.get("Guardian Item Number"):
                messagebox.showerror("Error", "Guardian Item Number is required.")
                return
            if selected_type == "Mobile Device":
                data["BFU/AFU"] = mobile_state.get()
                data["Airplane Mode"] = mobile_airplane.get()
                data["Micro SD Present"] = mobile_sd.get()
                data["Make"] = mobile_make.get()
                data["Model"] = mobile_model.get()
                data["SIM Present"] = mobile_sim.get()
                data["ICCID"] = mobile_iccid.get()
                data["Phone Number"] = mobile_phone.get()
                data["Serial Number"] = mobile_serial.get()
                data["IMEI"] = mobile_imei.get()
                data["Acquisition Type"] = mobile_acq.get()
            elif selected_type == "Hard Drive":
                data["Brand"] = hd_brand.get()
                data["Capacity"] = hd_capacity.get()
                data["Interface Type"] = hd_interface.get()
                data["Model"] = hd_model.get()
                data["Serial Number"] = hd_serial.get()
                data["Firmware Version"] = hd_firmware.get()
                data["Condition"] = hd_condition.get()
            elif selected_type == "SSD":
                data["Brand"] = ssd_brand.get()
                data["Capacity"] = ssd_capacity.get()
                data["Interface Type"] = ssd_interface.get()
                data["Model"] = ssd_model.get()
                data["Serial Number"] = ssd_serial.get()
                data["Firmware Version"] = ssd_firmware.get()
                data["Condition"] = ssd_condition.get()
            elif selected_type == "Drone":
                data["Manufacturer"] = drone_manufacturer.get()
                data["Model"] = drone_model.get()
                data["Serial Number"] = drone_serial.get()
                data["Firmware Version"] = drone_firmware.get()
                data["Battery Health"] = drone_battery.get()
                data["Camera Resolution"] = drone_camera.get()
                data["Acquisition Type"] = drone_acq.get()
            elif selected_type == "Infotainment System":
                data["Manufacturer"] = info_manuf.get()
                data["Model"] = info_model.get()
                data["Software Version"] = info_soft.get()
                data["Serial Number"] = info_serial.get()
                data["Firmware Version"] = info_firmware.get()
                data["Vehicle Make"] = info_vehicle_make.get()
                data["Vehicle Model"] = info_vehicle_model.get()
            elif selected_type == "Desktop Computer":
                data["Manufacturer"] = desk_manuf.get()
                data["Model"] = desk_model.get()
                data["Serial Number"] = desk_serial.get()
                data["Operating System"] = desk_os.get()
                data["CPU"] = desk_cpu.get()
                data["RAM"] = desk_ram.get()
                data["Storage Capacity"] = desk_storage.get()
                data["Acquisition Type"] = desk_acq.get()
            elif selected_type == "Laptop":
                data["Manufacturer"] = lap_manuf.get()
                data["Model"] = lap_model.get()
                data["Serial Number"] = lap_serial.get()
                data["Operating System"] = lap_os.get()
                data["CPU"] = lap_cpu.get()
                data["RAM"] = lap_ram.get()
                data["Storage Capacity"] = lap_storage.get()
                data["Battery Health"] = lap_battery.get()
            elif selected_type == "All Other":
                data["Device Description"] = other_desc.get("1.0", tk.END).strip()
                data["Make"] = other_make.get()
                data["Model"] = other_model.get()
                data["Serial Number"] = other_serial.get()
            if index is not None:
                self.devices[index] = data
                self.devices_listbox.delete(index)
                self.devices_listbox.insert(index, f"{data.get('Device Type')} - {data.get('Guardian Item Number')}")
            else:
                self.devices.append(data)
                self.devices_listbox.insert(tk.END, f"{data.get('Device Type')} - {data.get('Guardian Item Number')}")
            device_window.destroy()
        
        save_button = ttk.Button(device_window, text="Save Device", command=save_device)
        save_button.grid(row=3, column=0, columnspan=2, pady=10)
    
    def edit_device(self, event):
        selection = self.devices_listbox.curselection()
        if selection:
            index = selection[0]
            device_info = self.devices[index]
            self.add_device(device_info=device_info, index=index)
    
    def generate_report(self):
        if (not self.report_number.get() or not self.summary_text.get("1.0", 'end-1c') or 
            not self.drop_off_person.get() or not self.legal_basis.get() or 
            not any(var.get() for var in self.program_used_vars.values()) or 
            not any(var.get() for var in self.decoder_used_vars.values()) or 
            not self.md5_hash.get("1.0", 'end-1c') or not self.disposition.get()):
            messagebox.showerror("Error", "Please fill in all fields before generating the report.")
            return
        report_number = self.report_number.get()
        executive_summary = self.summary_text.get("1.0", 'end-1c')
        item_drop_off_date = self.drop_off_date.get()
        drop_off_person = self.drop_off_person.get()
        dropoff_date = self.dropoff_date.get()
        legal_basis = self.legal_basis.get()
        program_used = ", ".join([option for option, var in self.program_used_vars.items() if var.get()])
        decoder_used = ", ".join([option for option, var in self.decoder_used_vars.items() if var.get()])
        md5_hash_values = self.md5_hash.get("1.0", 'end-1c')
        disposition_date = self.disposition_date.get()
        disposition = self.disposition.get()
        template = f"""
Digital Forensics Report
Report Number: {report_number}

Introduction:

This report contains information about digital devices provided to the Raleigh Police Department's Digital Forensics Lab for the crime referenced in the master report, including details on the acquisition and decoding of the digital items processed by the lab. The report may only partially analyze some of the digital evidence obtained. Any analysis performed will be annotated with the header "Analysis" and/or "Executive Summary," depending on the specificity of the investigator's request.

Executive Summary:
  
{executive_summary}

Acquisition:

On {item_drop_off_date}, {drop_off_person} submitted a request for a digital forensic examination on the listed digital device(s).

The listed device(s) were brought to the lab on {dropoff_date}. The state of each device upon arrival is listed next to it below.

If a device arrived at the lab powered on after being confiscated, it is considered HOT; if it arrived powered off, it is considered COLD. If unknown, it is listed as UNKNOWN.

Legal Authority: {legal_basis}

Device(s):
"""
        for device in self.devices:
            template += "\n----------------------------------------\n"
            template += f"Device Type: {device.get('Device Type', 'N/A')}\n"
            for key, value in device.items():
                if key != "Device Type":
                    template += f"{key}: {value}\n"
            template += "----------------------------------------\n"
        
        template += f"""
Methodology:

The device(s) was/were acquired using {program_used}.

The listed item(s) were connected to a forensic workstation computer, and the extraction method was chosen accordingly.

The extracted data was then loaded into {decoder_used}.

The analysis software decodes the raw extracted data, allowing a trained digital forensic examiner to view the information in a human-readable format.

All extracted data was exported and saved in a separate folder in a report format.

The following hash values were obtained for the extracted data after being placed in a .zip container.

This protects the data and provides a long-term storage method in our digital evidence management system (DEMS).

The data was hashed using the MD5 algorithm, which provides a unique 32-hexadecimal value for the file.

MD5 Hash Values:

{md5_hash_values}

The Raleigh Police Department uses Cellebrite's Guardian solution to store and maintain digital evidence in the cloud.
The above files and their hash values are stored in the digital evidence management system.

On {disposition_date}, the device(s) was/were {disposition}.

The digital evidence obtained was placed into our digital evidence management system, where the lead investigator or requesting officer is provided an electronic copy.

Forensic Tools Used: 

{program_used}, {decoder_used}

Definitions:

Hash Value: A unique code generated when a file or data is processed through a mathematical algorithm, similar to a digital fingerprint.

Physical Extraction: The process of obtaining a copy of every binary data in a device's flash memory, used when traditional methods are not possible or the device is damaged.

Full File System Extraction: A complete copy of all the data stored on a device, including file organization information.

Advanced Logical Extraction: A method to obtain a copy of selected data (e.g., pictures, contacts, messages) that is directly accessible to the user.

SIM Data Extraction: A process to obtain data from the SIM card, which is a small, removable smart card used in mobile devices that contains information such as contacts, text messages, and network authentication details.

Hard Drive: The primary storage device of a computer, which stores all software programs and data.

Solid State Drive (SSD): A storage device that uses integrated circuit assemblies to store data persistently.

Flash Memory: A type of non-volatile memory that can be electrically erased and reprogrammed.

End of Report.
"""
        report_window = tk.Toplevel(self.root)
        report_window.title("Generated Report")
        report_text = scrolledtext.ScrolledText(report_window, width=100, height=40)
        report_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        report_text.insert("1.0", template)
        report_text.configure(state='disabled')
        
        def export_report():
            file_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                     filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")],
                                                     title="Save Report As")
            if file_path:
                try:
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.write(template)
                    messagebox.showinfo("Export Success", f"Report successfully exported to:\n{file_path}")
                except Exception as e:
                    messagebox.showerror("Export Error", f"An error occurred while exporting the report:\n{e}")
        
        export_button = ttk.Button(report_window, text="Export to .txt", command=export_report)
        export_button.pack(pady=5)

if __name__ == "__main__":
    root = tk.Tk()
    app = ReportApp(root)
    root.mainloop()
