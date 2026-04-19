import threading
import queue
import json
import os
import time
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import serial
import serial.tools.list_ports
import mido
import ttkbootstrap as tb
from ttkbootstrap.constants import *

DEFAULT_CONFIG_FILE = "mappings.json"

# UI 统一列宽，用于强制对齐映射表
COL_MSG_W = 20
COL_CC_W = 10
COL_VAL_W = 10
COL_ACT_W = 10

class SerialMidiControllerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Serial -> MIDI CC Controller")
        self.root.geometry("1000x700")

        self.style = tb.Style()

        # Variables
        self.serial_port = tk.StringVar()
        self.baud_rate = tk.StringVar(value="115200")
        self.midi_port = tk.StringVar()
        self.is_running = False

        # Floating Window Variables
        self.float_win = None
        self.pedal_status_var = tk.StringVar(value="Waiting...")
        self.last_received_var = tk.StringVar(value="-")

        self.mapping_rows = []
        self.worker_thread = None
        self.stop_event = threading.Event()
        self.line_queue = queue.Queue()

        self.serial_obj = None
        self.midi_out = None

        self._build_ui()
        self._refresh_ports()
        self._load_default_mappings()

    # ---------------------- UI 构建 ----------------------
    def _build_ui(self):
        main = ttk.Frame(self.root)
        main.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        # ====== 顶部设置区域 ======
        top_frame = ttk.Frame(main)
        top_frame.pack(fill=tk.X, pady=(0, 10))

        ser_frame = ttk.LabelFrame(top_frame, text="Serial Settings", padding=8)
        ser_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        ttk.Label(ser_frame, text="Port:").grid(row=0, column=0, sticky=tk.W, padx=2)
        self.ser_combo = ttk.Combobox(ser_frame, textvariable=self.serial_port, width=15)
        self.ser_combo.grid(row=0, column=1, padx=(2, 10))
        ttk.Label(ser_frame, text="Baud:").grid(row=0, column=2, sticky=tk.W, padx=2)
        self.baud_entry = ttk.Entry(ser_frame, textvariable=self.baud_rate, width=8)
        self.baud_entry.grid(row=0, column=3, padx=(2, 10))
        tb.Button(ser_frame, text="Refresh", command=self._refresh_ports).grid(row=0, column=4, padx=2)

        midi_frame = ttk.LabelFrame(top_frame, text="MIDI Settings", padding=8)
        midi_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        ttk.Label(midi_frame, text="MIDI Out:").grid(row=0, column=0, sticky=tk.W, padx=2)
        self.midi_combo = ttk.Combobox(midi_frame, textvariable=self.midi_port, width=20)
        self.midi_combo.grid(row=0, column=1, padx=(2, 10))
        tb.Button(midi_frame, text="Refresh", command=self._refresh_midi_ports).grid(row=0, column=2, padx=2)

        ctrl_frame = ttk.Frame(top_frame)
        ctrl_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
        self.start_btn = tb.Button(ctrl_frame, text="Start Processing", bootstyle="success", command=self.toggle_start)
        self.start_btn.pack(side=tk.TOP, fill=tk.X, pady=(5, 5))
        self.float_btn = tb.Button(ctrl_frame, text="Show Pedal Status Window", bootstyle="info-outline", command=self.toggle_float_window)
        self.float_btn.pack(side=tk.TOP, fill=tk.X)

        # ====== 中部映射区域 ======
        mapping_frame = ttk.LabelFrame(main, text="Mappings (Serial -> CC#, Value)")
        mapping_frame.pack(fill=tk.BOTH, expand=False, pady=(0, 10))

        btn_row = ttk.Frame(mapping_frame)
        btn_row.pack(fill=tk.X, padx=6, pady=6)
        tb.Button(btn_row, text="+ Add Mapping", command=self.add_mapping_row).pack(side=tk.LEFT)
        tb.Button(btn_row, text="Load...", bootstyle="secondary", command=self.load_mappings).pack(side=tk.LEFT, padx=6)
        tb.Button(btn_row, text="Save...", bootstyle="secondary", command=self.save_mappings).pack(side=tk.LEFT)

        header = ttk.Frame(mapping_frame)
        header.pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Label(header, text="Serial Message", width=COL_MSG_W, font=("", 10, "bold")).pack(side=tk.LEFT, padx=5)
        ttk.Label(header, text="CC #", width=COL_CC_W, font=("", 10, "bold")).pack(side=tk.LEFT, padx=5)
        ttk.Label(header, text="Value", width=COL_VAL_W, font=("", 10, "bold")).pack(side=tk.LEFT, padx=5)
        ttk.Label(header, text="Action", width=COL_ACT_W, font=("", 10, "bold")).pack(side=tk.LEFT, padx=5)

        self.rows_container = ttk.Frame(mapping_frame)
        self.rows_container.pack(fill=tk.X, padx=6, pady=(0, 6))

        # ====== 底部日志 ======
        lower = ttk.Frame(main)
        lower.pack(fill=tk.BOTH, expand=True)

        left_log = ttk.LabelFrame(lower, text="Event Log")
        left_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,6))
        self.log_text = tk.Text(left_log, height=12, state=tk.DISABLED, bg="#1e1e1e", fg="#00ff00")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        right_panel = ttk.LabelFrame(lower, text="Live Monitor")
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, ipadx=6, ipady=6)
        ttk.Label(right_panel, text="Last Received:").pack(anchor=tk.W, pady=(5,0))
        ttk.Label(right_panel, textvariable=self.last_received_var, font=("Consolas", 16, "bold"), bootstyle="primary").pack(anchor=tk.W, pady=(0,15))
        tb.Button(right_panel, text="About App", bootstyle="link", command=self.show_about).pack(side=tk.BOTTOM, pady=6)

    # ---------------------- 悬浮窗显示 ----------------------
    def toggle_float_window(self):
        if self.float_win is None or not self.float_win.winfo_exists():
            self._create_float_window()
        else:
            self._close_float_window()

    def _create_float_window(self):
        self.float_win = tk.Toplevel(self.root)
        self.float_win.title("Pedal Status")
        self.float_win.geometry("260x120")
        self.float_win.attributes("-topmost", True)
        self.float_win.protocol("WM_DELETE_WINDOW", self._close_float_window)

        lbl = ttk.Label(self.float_win, textvariable=self.pedal_status_var, font=("Helvetica", 24, "bold"), anchor="center", justify="center")
        lbl.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)
        self.float_btn.configure(text="Hide Pedal Status Window", bootstyle="danger-outline")

    def _close_float_window(self):
        if self.float_win:
            self.float_win.destroy()
            self.float_win = None
        self.float_btn.configure(text="Show Pedal Status Window", bootstyle="info-outline")

    def _parse_pedal_status(self, line):
        if line.endswith("B"): return f"🟢 {line[:-1]}\nPRESSED"
        elif line.endswith("A"): return f"⚪ {line[:-1]}\nRELEASED"
        else: return f"💬\n{line}"

    # ---------------------- 核心操作 ----------------------
    def add_mapping_row(self, msg_text="", cc_num="64", cc_val="127"):
        row = ttk.Frame(self.rows_container)
        row.pack(fill=tk.X, pady=2)
        
        msg_var = tk.StringVar(value=msg_text)
        cc_var = tk.StringVar(value=str(cc_num))
        val_var = tk.StringVar(value=str(cc_val))

        ttk.Entry(row, textvariable=msg_var, width=COL_MSG_W).pack(side=tk.LEFT, padx=5)
        ttk.Entry(row, textvariable=cc_var, width=COL_CC_W).pack(side=tk.LEFT, padx=5)
        ttk.Entry(row, textvariable=val_var, width=COL_VAL_W).pack(side=tk.LEFT, padx=5)
        
        del_btn = tb.Button(row, text="Delete", bootstyle="danger-outline", width=COL_ACT_W, 
                            command=lambda: self._delete_row(row, (msg_var, cc_var, val_var)))
        del_btn.pack(side=tk.LEFT, padx=5)

        self.mapping_rows.append((msg_var, cc_var, val_var))

    def _delete_row(self, frame, vars_tuple):
        if vars_tuple in self.mapping_rows:
            self.mapping_rows.remove(vars_tuple)
        frame.destroy()

    def _refresh_ports(self):
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.ser_combo['values'] = ports
        if ports and not self.serial_port.get():
            self.serial_port.set(ports[0])
        self._refresh_midi_ports()

    def _refresh_midi_ports(self):
        try: names = mido.get_output_names()
        except Exception: names = []
        self.midi_combo['values'] = names
        if names and not self.midi_port.get():
            self.midi_port.set(names[0])

    def show_about(self):
        messagebox.showinfo("About", "Serial -> MIDI CC Controller\nOptimized for ultra-low CPU usage.")

    def save_mappings(self):
        data = []
        for msg_var, cc_var, val_var in self.mapping_rows:
            msg = msg_var.get().strip()
            if not msg: continue
            try: data.append({"msg": msg, "cc": int(cc_var.get()), "val": int(val_var.get())})
            except ValueError:
                messagebox.showwarning("Invalid mapping", f"CC and Value must be integers for message '{msg}'")
                return
        file = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files","*.json"), ("All","*")])
        if file:
            with open(file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)

    def load_mappings(self):
        file = filedialog.askopenfilename(filetypes=[("JSON files","*.json"), ("All","*")])
        if not file: return
        try:
            with open(file, 'r', encoding='utf-8') as f: data = json.load(f)
            for child in list(self.rows_container.winfo_children()): child.destroy()
            self.mapping_rows.clear()
            for item in data: self.add_mapping_row(item.get('msg',''), item.get('cc',64), item.get('val',127))
        except Exception as e: messagebox.showerror("Error", f"Failed to load JSON: {e}")

    def _load_default_mappings(self):
        if os.path.exists(DEFAULT_CONFIG_FILE):
            try:
                with open(DEFAULT_CONFIG_FILE, 'r', encoding='utf-8') as f:
                    for item in json.load(f): self.add_mapping_row(item.get('msg',''), item.get('cc',64), item.get('val',127))
                return
            except Exception: pass
            
        # 恢复了完整的默认配置 P15 ~ P18
        defaults = [
            ("P15B", 64, 127), ("P15A", 64, 0),
            ("P16B", 66, 127), ("P16A", 66, 0),
            ("P17B", 67, 127), ("P17A", 67, 0),
            ("P18B", 68, 127), ("P18A", 68, 0),
        ]
        for msg, cc, val in defaults:
            self.add_mapping_row(msg, cc, val)

    def _log(self, text):
        ts = datetime.now().strftime('%H:%M:%S')
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{ts}] {text}\n")
        # 为防止日志无限堆积吃内存，超过 500 行自动清理一半
        if float(self.log_text.index('end-1c').split('.')[0]) > 500:
            self.log_text.delete('1.0', '250.0')
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def toggle_start(self):
        if self.is_running: self.stop()
        else: self.start()

    def start(self):
        port, baud, midi_name = self.serial_port.get().strip(), self.baud_rate.get().strip(), self.midi_port.get().strip()
        if not port or not midi_name: return messagebox.showwarning("Missing", "Select Serial and MIDI ports.")
        try: baud_i = int(baud)
        except ValueError: return messagebox.showwarning("Invalid", "Baudrate must be a number.")

        self.mapping = {}
        for msg_var, cc_var, val_var in self.mapping_rows:
            msg = msg_var.get().strip()
            if not msg: continue
            try: self.mapping[msg] = (int(cc_var.get()), int(val_var.get()))
            except ValueError: return messagebox.showwarning("Error", f"Invalid values for '{msg}'")

        try:
            if self.midi_out: self.midi_out.close()
            self.midi_out = mido.open_output(midi_name)
        except Exception as e: return messagebox.showerror("MIDI Error", f"Failed: {e}")

        try:
            if self.serial_obj and self.serial_obj.is_open: self.serial_obj.close()
            self.serial_obj = serial.Serial(port, baud_i, timeout=1)
        except Exception as e: return messagebox.showerror("Serial Error", f"Failed: {e}")

        self.stop_event.clear()
        self.worker_thread = threading.Thread(target=self._worker_loop, daemon=True)
        self.worker_thread.start()
        
        self.is_running = True
        self.start_btn.configure(text="Stop Processing", bootstyle="danger")
        self._log("Process started...")
        self.root.after(100, self._poll_queue)

    def stop(self):
        self.stop_event.set()
        self.is_running = False
        self.start_btn.configure(text="Start Processing", bootstyle="success")
        try:
            if self.serial_obj and self.serial_obj.is_open: self.serial_obj.close()
            if self.midi_out: self.midi_out.close()
        except Exception: pass
        self._log("Process stopped.")

    def _worker_loop(self):
        ser, midi = self.serial_obj, self.midi_out
        while not self.stop_event.is_set():
            try: raw = ser.readline()
            except Exception as e:
                self._log(f"Read error: {e}")
                self.stop_event.set()
                break

            if not raw:
                # 配合 timeout=1，如果没有数据，短暂休眠释放 CPU
                time.sleep(0.01)
                continue

            try: line = raw.decode(errors='ignore').strip()
            except Exception: continue

            if not line: continue
            
            # 将消息扔给 UI 线程
            self.line_queue.put(line)

            # 瞬间处理 MIDI 映射，0 延迟
            mapped = self.mapping.get(line)
            if mapped:
                cc_num, cc_val = mapped
                try:
                    midi.send(mido.Message('control_change', control=cc_num, value=cc_val))
                    self._log(f"MIDI Sent: CC{cc_num} = {cc_val} ({line})")
                except Exception: pass

    def _poll_queue(self):
        latest_line = None
        # 【优化】一次性抽干队列，只取最后一个状态，避免连踩时发生无效的 UI 刷新
        try:
            while True:
                latest_line = self.line_queue.get_nowait()
        except queue.Empty:
            pass
        
        if latest_line is not None:
            # 【优化】只有状态真正改变了，才触发界面重绘
            if latest_line != self.last_received_var.get():
                self.last_received_var.set(latest_line)
                self.pedal_status_var.set(self._parse_pedal_status(latest_line))
                
        if self.is_running:
            # 【优化】100ms 刷新一次 UI，10FPS 足够肉眼观看且极为省电
            self.root.after(100, self._poll_queue)

if __name__ == '__main__':
    root = tb.Window(themename="darkly")
    app = SerialMidiControllerApp(root)
    root.mainloop()