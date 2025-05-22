import tkinter as tk
from tkinter import ttk, messagebox

def center_window(window, parent):
    """Center a window relative to its parent."""
    window.update_idletasks()
    parent_x = parent.winfo_rootx()
    parent_y = parent.winfo_rooty()
    parent_width = parent.winfo_width()
    parent_height = parent.winfo_height()
    x = parent_x + (parent_width - window.winfo_width()) // 2
    y = parent_y + (parent_height - window.winfo_height()) // 2
    window.geometry(f"+{x}+{y}")

def create_labeled_entry(parent, label_text, variable, row, col, width=40, state='normal'):
    """Create a labeled entry widget and return the entry."""
    ttk.Label(parent, text=label_text).grid(row=row, column=col, sticky='w', padx=5, pady=2)
    entry = ttk.Entry(parent, textvariable=variable, width=width, state=state)
    entry.grid(row=row, column=col+1, sticky='w', padx=5, pady=2)
    return entry

def load_categories(file_path, defaults):
    """Load categories from a JSON file or return defaults."""
    import json
    try:
        with open(file_path, 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return defaults
    except Exception as e:
        messagebox.showwarning("Warning", f"Failed to load categories from file: {str(e)}. Using defaults.")
        return defaults