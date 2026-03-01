#!/usr/bin/env python3
"""
better office - Word Processor, Spreadsheet, and Presentation Creator
A basic office application with multiple document types
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font, colorchooser
from tkinter import scrolledtext
import json
import csv
from datetime import datetime

class OfficeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("better office")
        self.root.geometry("1000x700")
        
        # Create menu bar
        self.create_menu()
        
        # Create notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Initialize different modules
        self.word_processor = WordProcessor(self.notebook)
        self.spreadsheet = Spreadsheet(self.notebook)
        self.presentation = Presentation(self.notebook)
        
        # Add tabs
        self.notebook.add(self.word_processor.frame, text="Word Processor")
        self.notebook.add(self.spreadsheet.frame, text="Spreadsheet")
        self.notebook.add(self.presentation.frame, text="Presentation")
        
        # Status bar
        self.status_bar = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self.new_file)
        file_menu.add_command(label="Open", command=self.open_file)
        file_menu.add_command(label="Save", command=self.save_file)
        file_menu.add_command(label="Save As", command=self.save_as_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def new_file(self):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0:
            self.word_processor.new_document()
        elif current_tab == 1:
            self.spreadsheet.new_spreadsheet()
        elif current_tab == 2:
            self.presentation.new_presentation()
        self.status_bar.config(text="New file created")
    
    def open_file(self):
        current_tab = self.notebook.index(self.notebook.select())
        filetypes = []
        
        if current_tab == 0:  # Word Processor
            filetypes = [("Text files", "*.txt"), ("All files", "*.*")]
        elif current_tab == 1:  # Spreadsheet
            filetypes = [("CSV files", "*.csv"), ("All files", "*.*")]
        elif current_tab == 2:  # Presentation
            filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
        
        filename = filedialog.askopenfilename(filetypes=filetypes)
        if filename:
            if current_tab == 0:
                self.word_processor.open_document(filename)
            elif current_tab == 1:
                self.spreadsheet.open_spreadsheet(filename)
            elif current_tab == 2:
                self.presentation.open_presentation(filename)
            self.status_bar.config(text=f"Opened: {filename}")
    
    def save_file(self):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab == 0 and self.word_processor.current_file:
            self.word_processor.save_document(self.word_processor.current_file)
            self.status_bar.config(text=f"Saved: {self.word_processor.current_file}")
        elif current_tab == 1 and self.spreadsheet.current_file:
            self.spreadsheet.save_spreadsheet(self.spreadsheet.current_file)
            self.status_bar.config(text=f"Saved: {self.spreadsheet.current_file}")
        elif current_tab == 2 and self.presentation.current_file:
            self.presentation.save_presentation(self.presentation.current_file)
            self.status_bar.config(text=f"Saved: {self.presentation.current_file}")
        else:
            self.save_as_file()
    
    def save_as_file(self):
        current_tab = self.notebook.index(self.notebook.select())
        filetypes = []
        default_ext = ""
        
        if current_tab == 0:  # Word Processor
            filetypes = [("Text files", "*.txt"), ("Rich Text", "*.rtf"), ("All files", "*.*")]
            default_ext = ".txt"
        elif current_tab == 1:  # Spreadsheet
            filetypes = [("CSV files", "*.csv"), ("Tab Separated", "*.tsv"), ("All files", "*.*")]
            default_ext = ".csv"
        elif current_tab == 2:  # Presentation
            filetypes = [("JSON files", "*.json"), ("All files", "*.*")]
            default_ext = ".json"
        
        filename = filedialog.asksaveasfilename(
            defaultextension=default_ext,
            filetypes=filetypes
        )
        
        if filename:
            if current_tab == 0:
                self.word_processor.save_document(filename)
                self.word_processor.current_file = filename
            elif current_tab == 1:
                self.spreadsheet.save_spreadsheet(filename)
                self.spreadsheet.current_file = filename
            elif current_tab == 2:
                self.presentation.save_presentation(filename)
                self.presentation.current_file = filename
            self.status_bar.config(text=f"Saved as: {filename}")
    
    def show_about(self):
        messagebox.showinfo("About", "Simple Office Suite v1.0\n\nA basic office application with:\n- Word Processor\n- Spreadsheet\n- Presentation Creator")


class WordProcessor:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self.current_file = None
        
        # Toolbar
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        # Font controls
        ttk.Label(toolbar, text="Font:").pack(side=tk.LEFT, padx=5)
        self.font_family = ttk.Combobox(toolbar, values=list(font.families()), width=15)
        self.font_family.set("Arial")
        self.font_family.pack(side=tk.LEFT, padx=5)
        self.font_family.bind("<<ComboboxSelected>>", self.change_font)
        
        ttk.Label(toolbar, text="Size:").pack(side=tk.LEFT, padx=5)
        self.font_size = ttk.Combobox(toolbar, values=[8, 10, 12, 14, 16, 18, 20, 24, 28, 32], width=5)
        self.font_size.set("12")
        self.font_size.pack(side=tk.LEFT, padx=5)
        self.font_size.bind("<<ComboboxSelected>>", self.change_font)
        
        # Text formatting buttons
        self.bold_btn = ttk.Button(toolbar, text="B", width=3, command=self.toggle_bold)
        self.bold_btn.pack(side=tk.LEFT, padx=2)
        
        self.italic_btn = ttk.Button(toolbar, text="I", width=3, command=self.toggle_italic)
        self.italic_btn.pack(side=tk.LEFT, padx=2)
        
        self.underline_btn = ttk.Button(toolbar, text="U", width=3, command=self.toggle_underline)
        self.underline_btn.pack(side=tk.LEFT, padx=2)
        
        # Color button
        self.color_btn = ttk.Button(toolbar, text="Color", command=self.change_color)
        self.color_btn.pack(side=tk.LEFT, padx=5)
        
        # Alignment buttons
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)
        self.align_left_btn = ttk.Button(toolbar, text="←", width=3, command=lambda: self.set_alignment("left"))
        self.align_left_btn.pack(side=tk.LEFT, padx=2)
        
        self.align_center_btn = ttk.Button(toolbar, text="↔", width=3, command=lambda: self.set_alignment("center"))
        self.align_center_btn.pack(side=tk.LEFT, padx=2)
        
        self.align_right_btn = ttk.Button(toolbar, text="→", width=3, command=lambda: self.set_alignment("right"))
        self.align_right_btn.pack(side=tk.LEFT, padx=2)
        
        # Text widget with scrollbar
        text_frame = ttk.Frame(self.frame)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        self.text = tk.Text(text_frame, wrap=tk.WORD, undo=True, font=("Arial", 12))
        scrollbar = ttk.Scrollbar(text_frame, command=self.text.yview)
        self.text.config(yscrollcommand=scrollbar.set)
        
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Word count
        self.word_count_label = ttk.Label(self.frame, text="Words: 0 | Characters: 0")
        self.word_count_label.pack(side=tk.BOTTOM, anchor=tk.E, padx=5, pady=5)
        
        self.text.bind("<KeyRelease>", self.update_word_count)
    
    def change_font(self, event=None):
        try:
            family = self.font_family.get()
            size = int(self.font_size.get())
            self.text.config(font=(family, size))
        except:
            pass
    
    def toggle_bold(self):
        current_tags = self.text.tag_names("sel.first")
        if "bold" in current_tags:
            self.text.tag_remove("bold", "sel.first", "sel.last")
        else:
            self.text.tag_add("bold", "sel.first", "sel.last")
            self.text.tag_config("bold", font=("Arial", 12, "bold"))
    
    def toggle_italic(self):
        current_tags = self.text.tag_names("sel.first")
        if "italic" in current_tags:
            self.text.tag_remove("italic", "sel.first", "sel.last")
        else:
            self.text.tag_add("italic", "sel.first", "sel.last")
            self.text.tag_config("italic", font=("Arial", 12, "italic"))
    
    def toggle_underline(self):
        current_tags = self.text.tag_names("sel.first")
        if "underline" in current_tags:
            self.text.tag_remove("underline", "sel.first", "sel.last")
        else:
            self.text.tag_add("underline", "sel.first", "sel.last")
            self.text.tag_config("underline", underline=True)
    
    def change_color(self):
        color = colorchooser.askcolor()[1]
        if color:
            try:
                self.text.tag_add("color", "sel.first", "sel.last")
                self.text.tag_config("color", foreground=color)
            except:
                pass
    
    def set_alignment(self, align):
        try:
            self.text.tag_add(align, "sel.first", "sel.last")
            self.text.tag_config(align, justify=align)
        except:
            pass
    
    def update_word_count(self, event=None):
        content = self.text.get("1.0", tk.END)
        words = len(content.split())
        chars = len(content.strip())
        self.word_count_label.config(text=f"Words: {words} | Characters: {chars}")
    
    def new_document(self):
        if messagebox.askyesno("New Document", "Clear current document?"):
            self.text.delete("1.0", tk.END)
            self.current_file = None
    
    def open_document(self, filename):
        try:
            with open(filename, 'r') as file:
                self.text.delete("1.0", tk.END)
                self.text.insert("1.0", file.read())
                self.current_file = filename
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")
    
    def save_document(self, filename):
        try:
            with open(filename, 'w') as file:
                file.write(self.text.get("1.0", tk.END))
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {e}")


class Spreadsheet:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self.current_file = None
        self.rows = 20
        self.cols = 10
        
        # Toolbar
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="Add Row", command=self.add_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Add Column", command=self.add_column).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Delete Row", command=self.delete_row).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Delete Column", command=self.delete_column).pack(side=tk.LEFT, padx=5)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)
        ttk.Button(toolbar, text="Calculate Sum", command=self.calculate_sum).pack(side=tk.LEFT, padx=5)
        
        # Create spreadsheet frame with scrollbars
        self.sheet_frame = ttk.Frame(self.frame)
        self.sheet_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Canvas for scrolling
        self.canvas = tk.Canvas(self.sheet_frame)
        v_scrollbar = ttk.Scrollbar(self.sheet_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(self.sheet_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Frame inside canvas
        self.grid_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.grid_frame, anchor=tk.NW)
        
        # Create grid
        self.cells = []
        self.create_grid()
        
        self.grid_frame.update_idletasks()
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
    
    def create_grid(self):
        # Column headers
        for col in range(self.cols):
            header = tk.Label(self.grid_frame, text=chr(65+col), bg="lightgray", relief=tk.RIDGE, width=10)
            header.grid(row=0, column=col+1, sticky="nsew")
        
        # Row headers and cells
        for row in range(self.rows):
            # Row header
            header = tk.Label(self.grid_frame, text=str(row+1), bg="lightgray", relief=tk.RIDGE, width=5)
            header.grid(row=row+1, column=0, sticky="nsew")
            
            # Cells
            row_cells = []
            for col in range(self.cols):
                cell = tk.Entry(self.grid_frame, width=10, relief=tk.RIDGE)
                cell.grid(row=row+1, column=col+1, sticky="nsew")
                row_cells.append(cell)
            self.cells.append(row_cells)
    
    def add_row(self):
        self.rows += 1
        row = self.rows - 1
        
        # Row header
        header = tk.Label(self.grid_frame, text=str(row+1), bg="lightgray", relief=tk.RIDGE, width=5)
        header.grid(row=row+1, column=0, sticky="nsew")
        
        # Cells
        row_cells = []
        for col in range(self.cols):
            cell = tk.Entry(self.grid_frame, width=10, relief=tk.RIDGE)
            cell.grid(row=row+1, column=col+1, sticky="nsew")
            row_cells.append(cell)
        self.cells.append(row_cells)
        
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
    
    def add_column(self):
        self.cols += 1
        col = self.cols - 1
        
        # Column header
        header = tk.Label(self.grid_frame, text=chr(65+col) if col < 26 else f"A{chr(65+col-26)}", 
                         bg="lightgray", relief=tk.RIDGE, width=10)
        header.grid(row=0, column=col+1, sticky="nsew")
        
        # Cells
        for row in range(self.rows):
            cell = tk.Entry(self.grid_frame, width=10, relief=tk.RIDGE)
            cell.grid(row=row+1, column=col+1, sticky="nsew")
            self.cells[row].append(cell)
        
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
    
    def delete_row(self):
        if self.rows > 1:
            for widget in self.grid_frame.grid_slaves(row=self.rows):
                widget.destroy()
            if self.cells:
                self.cells.pop()
            self.rows -= 1
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
    
    def delete_column(self):
        if self.cols > 1:
            for row in range(self.rows + 1):
                for widget in self.grid_frame.grid_slaves(row=row, column=self.cols):
                    widget.destroy()
            for row_cells in self.cells:
                if row_cells:
                    row_cells.pop()
            self.cols -= 1
            self.canvas.config(scrollregion=self.canvas.bbox("all"))
    
    def calculate_sum(self):
        total = 0
        count = 0
        for row_cells in self.cells:
            for cell in row_cells:
                try:
                    value = float(cell.get())
                    total += value
                    count += 1
                except:
                    pass
        if count > 0:
            messagebox.showinfo("Sum", f"Total: {total}\nAverage: {total/count:.2f}")
        else:
            messagebox.showwarning("Sum", "No numeric values found")
    
    def new_spreadsheet(self):
        if messagebox.askyesno("New Spreadsheet", "Clear current spreadsheet?"):
            for row_cells in self.cells:
                for cell in row_cells:
                    cell.delete(0, tk.END)
            self.current_file = None
    
    def open_spreadsheet(self, filename):
        try:
            with open(filename, 'r') as file:
                reader = csv.reader(file)
                for i, row in enumerate(reader):
                    if i >= self.rows:
                        break
                    for j, value in enumerate(row):
                        if j >= self.cols:
                            break
                        self.cells[i][j].delete(0, tk.END)
                        self.cells[i][j].insert(0, value)
                self.current_file = filename
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")
    
    def save_spreadsheet(self, filename):
        try:
            with open(filename, 'w', newline='') as file:
                writer = csv.writer(file)
                for row_cells in self.cells:
                    row_data = [cell.get() for cell in row_cells]
                    writer.writerow(row_data)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {e}")


class Presentation:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self.current_file = None
        self.slides = []
        self.current_slide = 0
        
        # Toolbar
        toolbar = ttk.Frame(self.frame)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        
        ttk.Button(toolbar, text="New Slide", command=self.add_slide).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Delete Slide", command=self.delete_slide).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="← Prev", command=self.prev_slide).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Next →", command=self.next_slide).pack(side=tk.LEFT, padx=5)
        
        self.slide_label = ttk.Label(toolbar, text="Slide 0 of 0")
        self.slide_label.pack(side=tk.LEFT, padx=20)
        
        # Main content area
        content = ttk.Frame(self.frame)
        content.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Slide editor
        editor_frame = ttk.LabelFrame(content, text="Slide Editor")
        editor_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        ttk.Label(editor_frame, text="Title:").pack(anchor=tk.W, padx=5, pady=5)
        self.title_entry = ttk.Entry(editor_frame, font=("Arial", 16))
        self.title_entry.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(editor_frame, text="Content:").pack(anchor=tk.W, padx=5, pady=5)
        self.content_text = scrolledtext.ScrolledText(editor_frame, height=15, font=("Arial", 12))
        self.content_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Slide list
        list_frame = ttk.LabelFrame(content, text="Slides", width=200)
        list_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=5, pady=5)
        list_frame.pack_propagate(False)
        
        self.slide_listbox = tk.Listbox(list_frame)
        self.slide_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.slide_listbox.bind("<<ListboxSelect>>", self.on_slide_select)
        
        # Initialize with one slide
        self.add_slide()
    
    def add_slide(self):
        slide = {"title": f"Slide {len(self.slides) + 1}", "content": ""}
        self.slides.append(slide)
        self.slide_listbox.insert(tk.END, slide["title"])
        self.current_slide = len(self.slides) - 1
        self.update_slide_display()
    
    def delete_slide(self):
        if len(self.slides) > 0:
            if messagebox.askyesno("Delete", "Delete current slide?"):
                self.slides.pop(self.current_slide)
                self.slide_listbox.delete(self.current_slide)
                if self.current_slide >= len(self.slides):
                    self.current_slide = max(0, len(self.slides) - 1)
                self.update_slide_display()
    
    def prev_slide(self):
        self.save_current_slide()
        if self.current_slide > 0:
            self.current_slide -= 1
            self.update_slide_display()
    
    def next_slide(self):
        self.save_current_slide()
        if self.current_slide < len(self.slides) - 1:
            self.current_slide += 1
            self.update_slide_display()
    
    def on_slide_select(self, event):
        selection = self.slide_listbox.curselection()
        if selection:
            self.save_current_slide()
            self.current_slide = selection[0]
            self.update_slide_display()
    
    def save_current_slide(self):
        if 0 <= self.current_slide < len(self.slides):
            self.slides[self.current_slide]["title"] = self.title_entry.get()
            self.slides[self.current_slide]["content"] = self.content_text.get("1.0", tk.END).strip()
            self.slide_listbox.delete(self.current_slide)
            self.slide_listbox.insert(self.current_slide, self.slides[self.current_slide]["title"])
    
    def update_slide_display(self):
        if 0 <= self.current_slide < len(self.slides):
            slide = self.slides[self.current_slide]
            self.title_entry.delete(0, tk.END)
            self.title_entry.insert(0, slide["title"])
            self.content_text.delete("1.0", tk.END)
            self.content_text.insert("1.0", slide["content"])
            self.slide_label.config(text=f"Slide {self.current_slide + 1} of {len(self.slides)}")
            self.slide_listbox.selection_clear(0, tk.END)
            self.slide_listbox.selection_set(self.current_slide)
        else:
            self.title_entry.delete(0, tk.END)
            self.content_text.delete("1.0", tk.END)
            self.slide_label.config(text="Slide 0 of 0")
    
    def new_presentation(self):
        if messagebox.askyesno("New Presentation", "Clear current presentation?"):
            self.slides.clear()
            self.slide_listbox.delete(0, tk.END)
            self.current_slide = 0
            self.add_slide()
            self.current_file = None
    
    def open_presentation(self, filename):
        try:
            with open(filename, 'r') as file:
                data = json.load(file)
                self.slides = data.get("slides", [])
                self.slide_listbox.delete(0, tk.END)
                for slide in self.slides:
                    self.slide_listbox.insert(tk.END, slide["title"])
                self.current_slide = 0
                self.update_slide_display()
                self.current_file = filename
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")
    
    def save_presentation(self, filename):
        try:
            self.save_current_slide()
            data = {"slides": self.slides}
            with open(filename, 'w') as file:
                json.dump(data, file, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Could not save file: {e}")


def main():
    root = tk.Tk()
    app = OfficeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
