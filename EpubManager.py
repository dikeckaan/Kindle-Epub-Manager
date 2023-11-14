import os
import random
import tkinter as tk
from tkinter import messagebox
import win32com.client as win32

# Fixed folder path
FOLDER_PATH = "books"

def list_epubs():
    return [file for file in os.listdir(FOLDER_PATH) if file.endswith('.epub')]

def search(text):
    results = [file for file in list_epubs() if text.lower() in file.lower()]
    results_list.delete(0, tk.END)
    for file in results:
        results_list.insert(tk.END, file)

def select_random_epub():
    epub_files = list_epubs()
    if not epub_files:
        messagebox.showinfo("Info", "No .epub file found in folder.")
        return
    selected_file = random.choice(epub_files)
    results_list.delete(0, tk.END)
    results_list.insert(tk.END, selected_file)

def open_book():
    selected_file = results_list.get(tk.ACTIVE)
    if not selected_file:
        return
    selected_file_path = os.path.join(FOLDER_PATH, selected_file)
    try:
        os.startfile(selected_file_path)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while opening the file: {e}")

def send_to_kindle():
    selected_file = results_list.get(tk.ACTIVE)
    if not selected_file:
        return
    selected_file_path = os.path.abspath(os.path.join(FOLDER_PATH, selected_file))

    kindle_email = "your_kindle_email@example.com"
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = kindle_email
        mail.Subject = "Convert"
        mail.Body = f"Book sent: {selected_file}"
        mail.Attachments.Add(selected_file_path)
        mail.Display(True)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while creating the email: {e}")


def create_gui():
    window = tk.Tk()
    window.title("Epub Search and Send")
    window.geometry("600x500")
    
    # Set custom icon
    icon_path = "lib.ico"  # Replace with the correct path to your 'lib.ico' file
    window.iconbitmap(icon_path)


    # Dark theme colors
    dark_bg = "#363636"
    dark_fg = "#ffffff"
    dark_button_bg = "#505050"
    dark_listbox_bg = "#202020"

    window.configure(bg=dark_bg)

    search_bar = tk.Entry(window, bg=dark_listbox_bg, fg=dark_fg)
    search_bar.pack(pady=10)
    search_bar.bind('<KeyRelease>', lambda event: search(search_bar.get()))

    random_button = tk.Button(window, text="Find Random Book", command=select_random_epub, bg=dark_button_bg, fg=dark_fg)
    random_button.pack(pady=10)

    global results_list
    results_list = tk.Listbox(window, width=50, height=15, bg=dark_listbox_bg, fg=dark_fg)
    results_list.pack(pady=10)

    open_button = tk.Button(window, text="Open Book", command=open_book, bg=dark_button_bg, fg=dark_fg)
    open_button.pack(pady=5)

    kindle_button = tk.Button(window, text="Send to Kindle", command=send_to_kindle, bg=dark_button_bg, fg=dark_fg)
    kindle_button.pack(pady=5)

    window.mainloop()

create_gui()
