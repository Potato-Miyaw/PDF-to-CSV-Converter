import tkinter as tk
from tkinter import filedialog
import aspose.pdf as ap
import os

# Global variables for color palette
light_mode = True
light_background_color = "#747474"
light_sub_background_color = "#B1B1B1"
light_text_color = "#0A0708"
light_entry_color = "#FFFFFF"
dark_background_color = "#1e1e1e"
dark_sub_background_color = "#252526"
dark_text_color = "#9AC5F4"
dark_entry_color = "#3e3e42"

# Function to toggle between light mode and dark mode
def toggle_mode():
    global light_mode
    light_mode = not light_mode
    if light_mode:
        # Light Mode
        window.configure(bg=light_background_color)
        input_pdf_label.configure(bg=light_background_color, fg=light_text_color)
        input_pdf_entry.configure(bg=light_entry_color, fg=light_text_color)
        input_pdf_button.configure(bg=light_sub_background_color, fg=light_text_color)
        output_csv_label.configure(bg=light_background_color, fg=light_text_color)
        output_csv_entry.configure(bg=light_entry_color, fg=light_text_color)
        output_csv_button.configure(bg=light_sub_background_color, fg=light_text_color)
        store_button.configure(bg=light_sub_background_color, fg=light_text_color)
        convert_button.configure(bg=light_sub_background_color, fg=light_text_color)
        log_text.configure(bg=light_entry_color, fg=light_text_color)
    else:
        # Dark Mode
        window.configure(bg=dark_background_color)
        input_pdf_label.configure(bg=dark_background_color, fg=dark_text_color)
        input_pdf_entry.configure(bg=dark_entry_color, fg=dark_text_color)
        input_pdf_button.configure(bg=dark_sub_background_color, fg=dark_text_color)
        output_csv_label.configure(bg=dark_background_color, fg=dark_text_color)
        output_csv_entry.configure(bg=dark_entry_color, fg=dark_text_color)
        output_csv_button.configure(bg=dark_sub_background_color, fg=dark_text_color)
        store_button.configure(bg=dark_sub_background_color, fg=dark_text_color)
        convert_button.configure(bg=dark_sub_background_color, fg=dark_text_color)
        log_text.configure(bg=dark_entry_color, fg=dark_text_color)

# Converting .pdf to .csv
def convert_to_csv():
    input_pdf = input_pdf_entry.get()
    output_csv = output_csv_entry.get()

    if input_pdf and output_csv:
        document = ap.Document(input_pdf)
        save_option = ap.ExcelSaveOptions()
        save_option.format = ap.ExcelSaveOptions.ExcelFormat.CSV
        print(type(save_option))
        document.save(output_csv, save_option)

        # Extract file names
        pdf_name = os.path.basename(input_pdf)
        csv_name = os.path.basename(output_csv)

        log_text.insert(tk.END, f"Conversion completed successfully.\nPDF File: {pdf_name}\nCSV File: {csv_name}\n")
        log_text.see(tk.END)
    else:
        log_text.insert(tk.END, "Please provide both input and output file paths.\n")
        log_text.see(tk.END)

# Storing the CSV file in the same directory as the PDF with the same name
def store_csv_in_same_directory():
    input_pdf = input_pdf_entry.get()
    if input_pdf:
        input_dir = os.path.dirname(input_pdf)
        input_filename = os.path.basename(input_pdf)
        csv_filename = os.path.splitext(input_filename)[0] + ".csv"
        output_csv = os.path.join(input_dir, csv_filename)
        output_csv_entry.delete(0, tk.END)
        output_csv_entry.insert(0, output_csv)
        log_text.insert(tk.END, "CSV file will be stored in the same directory as the PDF.\n")
        log_text.see(tk.END)
    else:
        log_text.insert(tk.END, "Please provide the input PDF file.\n")
        log_text.see(tk.END)

# Browsing input file
def browse_input_pdf():
    input_pdf = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    input_pdf_entry.delete(0, tk.END)
    input_pdf_entry.insert(0, input_pdf)

# Browsing output file
def browse_output_csv():
    output_csv = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
    output_csv_entry.delete(0, tk.END)
    output_csv_entry.insert(0, output_csv)

# Create the main window
window = tk.Tk()
window.title("PDF to CSV Converter")
window.geometry("600x400")
window.configure(bg=light_background_color)

# Create menubar
menubar = tk.Menu(window)

# Create the "Mode" menu
mode_menu = tk.Menu(menubar, tearoff=0)

def set_light_mode():
    global light_mode
    if not light_mode:
        toggle_mode()

def set_dark_mode():
    global light_mode
    if light_mode:
        toggle_mode()

mode_menu.add_command(label="Light Mode", command=set_light_mode, background=light_sub_background_color, foreground=light_text_color)
mode_menu.add_command(label="Dark Mode", command=set_dark_mode, background=dark_sub_background_color, foreground=dark_text_color)

menubar.add_cascade(label="Mode", menu=mode_menu)

# Configure the window's menu
window.config(menu=menubar)

# Create and configure a grid for the window
window.grid_columnconfigure(1, weight=1)
window.grid_rowconfigure(4, weight=1)

# Create and place input file path label and entry
input_pdf_label = tk.Label(window, text="Input PDF:", bg=light_background_color, fg=light_text_color)
input_pdf_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")
input_pdf_entry = tk.Entry(window)
input_pdf_entry.grid(row=0, column=1, padx=10, pady=10, sticky="we")
input_pdf_button = tk.Button(window, text="Browse", command=browse_input_pdf, bg=light_sub_background_color, fg=light_text_color)
input_pdf_button.grid(row=0, column=2, padx=10, pady=10)

# Create and place output file path label and entry
output_csv_label = tk.Label(window, text="Output CSV:", bg=light_background_color, fg=light_text_color)
output_csv_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")
output_csv_entry = tk.Entry(window)
output_csv_entry.grid(row=1, column=1, padx=10, pady=10, sticky="we")
output_csv_button = tk.Button(window, text="Browse", command=browse_output_csv, bg=light_sub_background_color, fg=light_text_color)
output_csv_button.grid(row=1, column=2, padx=10, pady=10)

# Create and place store in same directory button
store_button = tk.Button(window, text="Store in Same Directory", command=store_csv_in_same_directory, bg=light_sub_background_color, fg=light_text_color)
store_button.grid(row=2, column=1, padx=10, pady=10, sticky="we")

# Create and place conversion button
convert_button = tk.Button(window, text="Convert to CSV", command=convert_to_csv, bg=light_sub_background_color, fg=light_text_color)
convert_button.grid(row=3, column=1, padx=10, pady=10, sticky="we")

# Create text box for logging
log_text = tk.Text(window)
log_text.grid(row=4, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

# Configure resizing behavior
window.grid_rowconfigure(4, weight=1)
window.grid_columnconfigure(0, weight=1)
window.grid_columnconfigure(1, weight=1)
window.grid_columnconfigure(2, weight=1)

# Start the main loop
window.mainloop()
