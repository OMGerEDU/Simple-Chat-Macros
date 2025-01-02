import keyboard
import time
from tkinter import *
from tkinter import messagebox, ttk
import sys
from PIL import Image
import pystray
import win32event
import win32api
import winerror
from threading import Thread
import os

mutex_name = "TextTyper_Mutex"
mutex = win32event.CreateMutex(None, 1, mutex_name)
if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
    messagebox.showerror("Error", "Application is already running!")
    sys.exit(1)

class TextTyperApp:
    def __init__(self):
        self.root = Tk()
        self.root.title("Text Typer")
        
        # Make window size dynamic based on screen size
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        window_width = int(screen_width * 0.6)
        window_height = int(screen_height * 0.8)
        self.root.geometry(f"{window_width}x{window_height}")
        
        self.running = True
        self.hotkeys = {}
        self.text_widgets = {}
        
        # Add bindings for paste and undo
        self.root.bind_class('Text', '<Control-v>', self.paste_text)
        self.root.bind_class('Text', '<Control-z>', self.undo_text)
        self.root.bind_class('Text', '<Control-y>', self.redo_text)
        self.root.bind_class('Text', '<Control-c>', self.copy_text)
        self.root.bind_class('Text', '<Control-x>', self.cut_text)
        
        self.load_config()
        self.create_gui()
        self.setup_tray()
        self.setup_hotkeys()

    def load_config(self):
        try:
            with open('config.txt', 'r', encoding='utf-8') as f:
                current_key = None
                current_text = []
                
                for line in f:
                    if '@@' in line:
                        if current_key and current_text:
                            self.hotkeys[current_key] = '\n'.join(current_text)
                        
                        current_key, text = line.strip().split('@@', 1)
                        current_text = [text] if text else []
                    else:
                        if current_key and line.strip():
                            current_text.append(line.strip())
                
                if current_key and current_text:
                    self.hotkeys[current_key] = '\n'.join(current_text)
                    
        except FileNotFoundError:
            # Default values
            self.hotkeys = {
                # Regular F-keys
                'F1': "ערב טוב כאן עומר מתכנית עמית במה אוכל לעזור?",
                'F2': "היי,  האם קיבלת SMS שמודיע כי נוספת למערכת?",
                'F4': ".אם עשית מעל 30 ימי מילואים, אתה זכאי לתכנית.\nעם זאת, טרם נפתח חשבון  לכל הזכאים.\nבחודש הקרוב לכל הזכאים הנותרים יקבלו SMS שמודיע להם כי חשבונתם התווסף למערכת.",
                'F6': "",
                'F7': "",
                'F8': "משהו נוסף שאוכל לעזור?",
                'F9': "",
                'F10': "המשך יום מקסים!",
                # Shift+F-keys
                'Shift+F1': "",
                'Shift+F2': "",
                'Shift+F3': "",
                'Shift+F4': "",
                'Shift+F5': "",
                'Shift+F6': "",
                # Ctrl+Numbers
                'Ctrl+0': "",
                'Ctrl+1': "",
                'Ctrl+2': "",
                'Ctrl+3': "",
                'Ctrl+4': "",
                'Ctrl+5': "",
                'Ctrl+6': "",
                'Ctrl+7': "",
                'Ctrl+8': "",
                'Ctrl+9': ""
            }
            self.save_config()

    def save_config(self):
        with open('config.txt', 'w', encoding='utf-8') as f:
            for key, text in self.hotkeys.items():
                f.write(f"{key}@@{text}\n")

    def create_gui(self):
        # Create toolbar
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=X, padx=5, pady=5)
        
        save_button = ttk.Button(toolbar, text="Save", command=self.save_configuration)
        save_button.pack(side=LEFT, padx=5)
        
        paste_button = ttk.Button(toolbar, text="Paste", command=self.paste_to_active)
        paste_button.pack(side=LEFT, padx=5)
        
        undo_button = ttk.Button(toolbar, text="Undo", command=self.undo_active)
        undo_button.pack(side=LEFT, padx=5)

        redo_button = ttk.Button(toolbar, text="Redo", command=self.redo_active)
        redo_button.pack(side=LEFT, padx=5)

        # Create status bar
        self.status_bar = ttk.Label(self.root, text="Ready", anchor=W)
        self.status_bar.pack(side=BOTTOM, fill=X, padx=5, pady=2)

        # Create main frame with notebook
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=BOTH, expand=True, padx=10, pady=5)

        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=BOTH, expand=True)

        # Regular F-keys tab
        regular_frame = ttk.Frame(notebook)
        notebook.add(regular_frame, text='F-keys')
        self.create_scrollable_frame(regular_frame, ['F1', 'F2', 'F4', 'F6', 'F7', 'F8', 'F9', 'F10'])

        # Shift+F-keys tab
        shift_f_frame = ttk.Frame(notebook)
        notebook.add(shift_f_frame, text='Shift + F-keys')
        self.create_scrollable_frame(shift_f_frame, ['Shift+F1', 'Shift+F2', 'Shift+F3', 'Shift+F4', 'Shift+F5', 'Shift+F6'])

        # Ctrl+Numbers tab
        ctrl_num_frame = ttk.Frame(notebook)
        notebook.add(ctrl_num_frame, text='Ctrl + Numbers')
        self.create_scrollable_frame(ctrl_num_frame, [f'Ctrl+{i}' for i in range(10)])

    def create_scrollable_frame(self, parent, keys):
        canvas = Canvas(parent)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollable_frame.grid_columnconfigure(1, weight=1)

        row = 0
        for key in keys:
            ttk.Label(scrollable_frame, text=f"{key}:").grid(row=row, column=0, padx=5, pady=5, sticky="nw")
            
            text_widget = Text(scrollable_frame, height=4, width=60, undo=True, maxundo=-1)
            text_widget.grid(row=row, column=1, padx=5, pady=5, sticky="ew")
            text_widget.insert('1.0', self.hotkeys.get(key, ""))
            
            text_widget.bind('<FocusIn>', lambda e, k=key: self.update_status(f"Editing {k}"))
            text_widget.bind('<FocusOut>', lambda e: self.update_status("Ready"))
            
            self.add_right_click_menu(text_widget)
            
            self.text_widgets[key] = text_widget
            
            ttk.Button(scrollable_frame, text="Test", 
                      command=lambda k=key: self.test_hotkey(k)).grid(row=row, column=2, padx=5, pady=5)
            
            row += 1

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def update_status(self, message):
        self.status_bar.config(text=message)

    def add_right_click_menu(self, widget):
        menu = Menu(self.root, tearoff=0)
        menu.add_command(label="Cut", command=lambda: self.cut_text(None, widget))
        menu.add_command(label="Copy", command=lambda: self.copy_text(None, widget))
        menu.add_command(label="Paste", command=lambda: self.paste_text(None, widget))
        menu.add_separator()
        menu.add_command(label="Undo", command=lambda: self.undo_text(None, widget))
        menu.add_command(label="Redo", command=lambda: self.redo_text(None, widget))
        
        def show_menu(event):
            menu.tk_popup(event.x_root, event.y_root)
        
        widget.bind('<Button-3>', show_menu)

    def cut_text(self, event, widget=None):
        try:
            widget = widget or event.widget
            if widget.tag_ranges(SEL):
                self.copy_text(event, widget)
                widget.delete(SEL_FIRST, SEL_LAST)
            return "break"
        except:
            pass

    def copy_text(self, event, widget=None):
        try:
            widget = widget or event.widget
            if widget.tag_ranges(SEL):
                selected_text = widget.get(SEL_FIRST, SEL_LAST)
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
            return "break"
        except:
            pass

    def paste_text(self, event, widget=None):
        try:
            widget = widget or event.widget
            clipboard_text = self.root.clipboard_get()
            widget.insert(INSERT, clipboard_text)
            return "break"
        except:
            pass

    def undo_text(self, event, widget=None):
        try:
            widget = widget or event.widget
            widget.edit_undo()
            return "break"
        except:
            pass

    def redo_text(self, event, widget=None):
        try:
            widget = widget or event.widget
            widget.edit_redo()
            return "break"
        except:
            pass

    def paste_to_active(self):
        focused_widget = self.root.focus_get()
        if isinstance(focused_widget, Text):
            self.paste_text(None, focused_widget)

    def undo_active(self):
        focused_widget = self.root.focus_get()
        if isinstance(focused_widget, Text):
            self.undo_text(None, focused_widget)

    def redo_active(self):
        focused_widget = self.root.focus_get()
        if isinstance(focused_widget, Text):
            self.redo_text(None, focused_widget)

    def test_hotkey(self, key):
        text = self.text_widgets[key].get('1.0', 'end-1c')
        type_string(text)

    def save_configuration(self):
        for key in self.text_widgets:
            self.hotkeys[key] = self.text_widgets[key].get('1.0', 'end-1c')
        self.save_config()
        self.setup_hotkeys()
        messagebox.showinfo("Success", "Configuration saved successfully!")

    def quit_application(self):
        self.save_configuration()
        self.running = False
        keyboard.unhook_all()
        self.icon.stop()
        self.root.quit()
        sys.exit(0)

    def setup_tray(self):
        image = Image.new('RGB', (64, 64), color='blue')
        menu = (pystray.MenuItem('Show Config', self.show_window),
                pystray.MenuItem('Exit', self.quit_application))
        self.icon = pystray.Icon("TextTyper", image, "Text Typer", menu)
        Thread(target=self.icon.run).start()

    def show_window(self):
        self.root.deiconify()

    def setup_hotkeys(self):
        keyboard.unhook_all()
        for key, text in self.hotkeys.items():
            if key.startswith('Shift+'):  # Shift+F keys
                _, fkey = key.split('+')
                keyboard.add_hotkey(f'shift+{fkey.lower()}', lambda t=text: type_string(t))
            elif key.startswith('Ctrl+'):  # Ctrl+Number keys
                _, num = key.split('+')
                keyboard.add_hotkey(f'ctrl+{num}', lambda t=text: type_string(t))
            else:  # Regular F-keys
                keyboard.add_hotkey(key.lower(), lambda t=text: type_string(t))

    def run(self):
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        self.root.mainloop()

    def minimize_to_tray(self):
        self.root.withdraw()

def type_string(text):
    time.sleep(0.1)
    keyboard.write(text)

if __name__ == "__main__":
    app = TextTyperApp()
    app.run()