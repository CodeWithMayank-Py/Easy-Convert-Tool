import tkinter as tk
from PIL import Image, ImageTk
import customtkinter as ctk
from tkinter import messagebox, filedialog
import warnings
import os
from image_converter import convert_jpeg_to_png, convert_png_to_jpeg, convert_image_to_pdf, convert_pdf_to_image
from document_converter import convert_pdf_to_word, convert_word_to_pdf, convert_excel_to_pdf, convert_pdf_to_excel, \
    convert_pdf_to_ppt, convert_ppt_to_pdf
from file_format_converter import convert_text_to_pdf, convert_csv_to_excel, convert_excel_to_csv
import pygame
import json
import urllib.request
from pathlib import Path


# Ignore Warnings
# Filter out the specific warning message
warnings.filterwarnings("ignore", category=UserWarning,
                        message="CTkButton Warning: Given image is not CTkImage but <class 'PIL.ImageTk.PhotoImage'>.")

# ================================================================================================================== #
# Update Logic Code Here

version_file_url = "https://github.com/CodeWithMayank-Py/Easy-Convert-Tool/blob/main/version.json"
google_drive_exe_link = ""


# Code to Retrieve the latest version from the version file
def get_installed_version():
    try:
        from version import version
        return version
    except ImportError:
        return None


def get_latest_version():
    try:
        response = urllib.request.urlopen(version_file_url)
        data = json.load(response)
        latest_version = data["version"]
        return latest_version
    except:
        return None


def download_latest_executable():
    try:
        urllib.request.urlretrieve(google_drive_exe_link, "latest_exe.exe")
        print("Latest version downloaded successfully.")
    except:
        print("Failed to download the latest version")


# Implement the update check logic in your software
installed_version = get_installed_version()
latest_version = get_latest_version()

if installed_version and latest_version and installed_version < latest_version:
    print("An update is available! Downloading the latest version....")
    download_latest_executable()
else:
    print("Your software is up to date.")


# =================================================================================================================== #


def clear_right_frame():
    # Remove all widgets from the right frame
    for widget in right_frame.winfo_children():
        widget.destroy()


def go_to_home():
    clear_right_frame()

    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Labels for bottom
    hm_powered_by = "CortexX"
    hm_disclaimer = "This tool is for personal use only."

    label = tk.Label(right_frame, text=f"Welcome \n\nto \n\nEasy Convert Tool",
                     font=("Arial", 50, "bold italic underline"), padx=20,
                     pady=20,
                     bg=app.cget("background"), fg="Lightcyan2")
    label.pack(fill=tk.BOTH, expand=True)

    # Create a label widget for the "Powered by" text
    hm_powered_by_label = tk.Label(right_frame, text=f"Powered by {hm_powered_by}", font=("Courier New", 24, "italic"),
                                   bg=app.cget("background"), fg="green2")
    hm_powered_by_label.pack()

    # Create a label widget for the disclaimer
    hm_disclaimer_label = tk.Label(right_frame, text=f"Disclaimer: {hm_disclaimer}",
                                   font=("Trebuchet MS", 14, "italic"),
                                   bg=app.cget("background"), fg="red3")
    hm_disclaimer_label.pack()



# Button 1: PNG ==> JPEG

def show_png_to_jpeg():
    # Clear the right frame
    clear_right_frame()

    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PNG to JPEG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4",
                         highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your PNG file here...", font=("Arial", 16), padx=20, pady=20,
                          bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("PNG Files", "*.png")])
        if file_path:
            if not file_path.lower().endswith('.png'):
                messagebox.showerror("Error", "Selected file is not in PNG format.")
            else:
                # Perform the necessary processing on the uploaded file
                # process_uploaded_file(file_path)
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path), font=("Arial", 14),
                                          padx=20, pady=10, bg=app.cget("background"), fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the PNG file to JPEG format
                jpeg_path = os.path.splitext(file_path)[0] + ".jpeg"
                success = convert_png_to_jpeg(file_path, jpeg_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully!", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".jpeg",
                                                                 filetypes=[("JPEG Files", "*.jpeg")])
                        if save_path:
                            os.rename(jpeg_path, save_path)
                            messagebox.showinfo("Download Complete", "JPEG file downloaded successfully.")

                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10, pady=5,
                                                bg=right_frame["bg"], fg="white", activebackground="green",
                                                highlightcolor="green", command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)




                else:
                    messagebox.showerror("Conversion Error", "Error converting PNG file to JPEG.")

    # Add the upload file button inside the box

    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 2 : JPEG ==> PNG


def show_jpeg_to_png():
    # Clear the right frame
    clear_right_frame()

    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PNG to JPEG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4",
                         highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your JPEG file here...", font=("Arial", 16), padx=20, pady=20,
                          bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("JPEG Files", "*.jpeg")])
        if file_path:
            if not file_path.lower().endswith('.jpeg'):
                messagebox.showerror("Error", "Selected file is not in JPEG format.")
            else:
                # Perform the necessary processing on the uploaded file
                # process_uploaded_file(file_path)
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path), font=("Arial", 14),
                                          padx=20, pady=10, bg=app.cget("background"), fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the JPEG file to PNG format
                png_path = os.path.splitext(file_path)[0] + ".png"
                success = convert_jpeg_to_png(file_path, png_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully!", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                                 filetypes=[("PNG Files", "*.png")])
                        if save_path:
                            os.rename(png_path, save_path)
                            messagebox.showinfo("Download Complete", "PNG file downloaded successfully.")

                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10, pady=5,
                                                bg=right_frame["bg"], fg="white", activebackground="green",
                                                highlightcolor="green", command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

                else:
                    messagebox.showerror("Conversion Error", "Error converting JPEG file to PNG.")

    # Add the upload file button inside the box
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 3 : JPEG \ PDF ==> PDF

def show_image_to_pdf():
    # Clear the right frame
    clear_right_frame()

    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for Image to PDF conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your JPEG or PNG file here...", font=("Arial", 16), padx=20, pady=20,
                          bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("JPEG Files", "*.jpeg"), ("PNG Files", "*.png")])
        if file_path:
            if not file_path.lower().endswith(('.jpeg', '.jpg', '.png')):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Perform the necessary processing on the uploaded file
                # process_uploaded_file(file_path)
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path), font=("Arial", 14),
                                          padx=20, pady=10, bg=app.cget("background"), fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the JPEG or PNG file to PDF format
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                success = convert_image_to_pdf(file_path, pdf_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                                 filetypes=[("PDF Files", "*.pdf")])
                        if save_path:
                            os.rename(pdf_path, save_path)
                            messagebox.showinfo("Download Complete", "PNG file downloaded successfully.")

                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10, pady=5,
                                                bg=right_frame["bg"], fg="white", activebackground="green",
                                                highlightcolor="green", command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Add the upload file button inside the box
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 4 : PDF to JPEG or PNG format

def show_pdf_to_png():
    # Clear the right frame
    clear_right_frame()

    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your PDF file here to convert to PNG....", font=("Arial", 16),
                          padx=20,
                          pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            if not file_path.lower().endswith('.pdf'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Perform the necessary processing on the uploaded file
                # process_uploaded_file(file_path)
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the PDF file to PNG format
                success = convert_pdf_to_image(file_path, '.png')

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".png",
                                                                 filetypes=[("PNG Files", "*.png")])

                        if save_path:
                            convert_pdf_to_image(file_path, '.png')
                            converted_file_path = os.path.splitext(file_path)[0] + "_1.png"
                            os.rename(converted_file_path, save_path)
                            messagebox.showinfo("Download Complete", "PNG file downloaded successfully.")

                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10, pady=5,
                                                bg=right_frame["bg"], fg="white", activebackground="green",
                                                highlightcolor="green", command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Add the upload file button inside the box
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Section 2: Document Parsing
# Frame 2 : Document Conversions
# Button 1: PDF to word

def show_pdf_to_word():
    # Clear the right frame
    clear_right_frame()

    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your PDF (.pdf) file here.....", font=("Arial", 16),
                          padx=20,
                          pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("WORD Files", "*.docx")])
        if file_path:
            if not file_path.lower().endswith('.docx'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Perform the necessary processing on the uploaded file
                # process_uploaded_file(file_path)
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"), fg="white")

                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the PDF file to WORD format
                word_path = os.path.splitext(file_path)[0] + ".docx"
                success = convert_pdf_to_word(file_path, word_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                                 filetypes=[("Word Files", "*.docx")])

                        if save_path:
                            os.rename(word_path, save_path)
                            messagebox.showinfo("Download Complete", "Word (.docx) file downloaded successfully.")
                        else:
                            messagebox.showerror("Save Error", "Unsupported file format.")
                            return

                        if not save_path:  # Check if save_path is empty
                            return

                        success = convert_pdf_to_word(file_path, save_path)

                        if success:
                            messagebox.showinfo("Download Complete", "Word file downloaded successfully.")

                    # Download button
                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10, pady=5,
                                                bg=right_frame["bg"], fg="white", activebackground="green",
                                                highlightcolor="green", command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Add the upload file button inside the box
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 2 : WORD to PDF format

def show_word_to_pdf():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your Word (.docx) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
        if file_path:
            if not file_path.lower().endswith('.docx'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the Word file to PDF format
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                success = convert_word_to_pdf(file_path, pdf_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                                 filetypes=[("PDF Files", "*.pdf")])
                        if save_path:
                            os.rename(pdf_path, save_path)
                            messagebox.showinfo("Download Complete", "WORD file downloaded successfully.")
                        else:
                            messagebox.showerror("Save Error", "Unsupported file format.")
                            return

                        if not save_path:  # Check if save_path is empty
                            return

                        success = convert_word_to_pdf(file_path, save_path)

                        if success:
                            messagebox.showinfo("Download Complete", "PDF file downloaded successfully.")

                    # Download button
                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                                pady=5, bg=right_frame["bg"], fg="white",
                                                activebackground="green", highlightcolor="green",
                                                command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Add the upload file button inside the box
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 3 : EXCEL to PDF format

def show_excel_to_pdf():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your Excel (.xlsx) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the Excel file to PDF format
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                success = convert_excel_to_pdf(file_path, pdf_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                                 filetypes=[("PDF Files", "*.pdf")])
                        if save_path:
                            os.rename(pdf_path, save_path)
                            messagebox.showinfo("Download Complete", "PDF file downloaded successfully.")
                        else:
                            messagebox.showerror("Save Error", "Unsupported file format.")
                            return

                        if not save_path:  # Check if save_path is empty
                            return

                        success = convert_excel_to_pdf(file_path, save_path)

                        if success:
                            messagebox.showinfo("Download Complete", "PDF file downloaded successfully.")

                    # Download button
                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                                pady=5, bg=right_frame["bg"], fg="white",
                                                activebackground="green", highlightcolor="green",
                                                command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Add the upload file button inside the box
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 4 : PDF to EXCEL Format

def show_pdf_to_excel():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your PDF (.pdf) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            if not file_path.lower().endswith('.pdf'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the PDF file to EXCEL format
                excel_path = os.path.splitext(file_path)[0] + ".xlsx"
                success = convert_pdf_to_excel(file_path, excel_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                                 filetypes=[("Excel Files", "*.xlsx")])
                        if save_path:
                            os.rename(excel_path, save_path)
                            messagebox.showinfo("Download Complete", "Excel (.xlsx) file downloaded successfully.")
                        else:
                            messagebox.showerror("Save Error", "Unsupported file format.")
                            return

                        if not save_path:  # Check if save_path is empty
                            return

                            success = convert_pdf_to_excel(file_path, save_path)

                            if success:
                                messagebox.showinfo("Download Complete", "EXCEl (.xlsx) file downloaded successfully.")

                    # Download button
                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                                pady=5, bg=right_frame["bg"], fg="white",
                                                activebackground="green", highlightcolor="green",
                                                command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Add the upload file button inside the box
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 5 : PDF to PowerPoint

def show_pdf_to_ppt():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your PDF (.pdf) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])

        if file_path:
            if not file_path.lower().endswith('.pdf'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the PDF file to PowerPoint format
                ppt_path = os.path.splitext(file_path)[0] + ".pptx"
                success = convert_pdf_to_ppt(file_path, ppt_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".pptx",
                                                                 filetypes=[("PowerPoint Files", "*.pptx")])

                        if save_path:
                            os.rename(ppt_path, save_path)
                            messagebox.showinfo("Download Complete", "PowerPoint (.pptx) file downloaded successfully.")
                        else:
                            messagebox.showerror("Save Error", "Unsupported file format.")
                            return

                        # Check if save_path is empty
                        if not save_path:
                            return

                        success = convert_pdf_to_ppt(file_path, save_path)

                        if success:
                            messagebox.showinfo("Download Complete", "PowerPoint (.pptx) file downloaded successfully.")

                    # Download button
                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                                pady=5, bg=right_frame["bg"], fg="white",
                                                activebackground="green", highlightcolor="green",
                                                command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Upload Button
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 6 : PowerPoint to PDF format

def show_ppt_to_pdf():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your PowerPoint (.pptx) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx")])
        if file_path:
            if not file_path.lower().endswith('.pptx'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")

            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the PowerPoint file to pdf format
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                success = convert_ppt_to_pdf(file_path, pdf_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully", font=("Arial", 14),
                                                padx=20, pady=10, bg=app.cget("background"), fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                    def download_file():
                        save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                                 filetypes=[("PDF Files", "*.pdf")])

                        if save_path:
                            os.rename(pdf_path, save_path)
                            messagebox.showinfo("Download Complete", "PDF (.pdf) file downloaded successfully.")
                        else:
                            messagebox.showerror("Save Error", "Unsupported file format.")
                            return

                        if not save_path:
                            return

                            success = convert_ppt_to_pdf(file_path, save_path)

                            if success:
                                messagebox.showinfo("Download Complete", "PDF (.pdf) file downloaded successfully.")

                    # Download Button

                    download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                                pady=5, bg=right_frame["bg"], fg="white",
                                                activebackground="green", highlightcolor="green",
                                                command=download_file)
                    download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Adding the upload file
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Section 3

# AUdio \ Video Conversation

# Button 1 : Mp3 to Mp4 format

def show_text_to_pdf():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your TEXT (.txt) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            if not file_path.lower().endswith('.txt'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the text file to pdf format
                pdf_path = os.path.splitext(file_path)[0] + ".pdf"
                success = convert_text_to_pdf(file_path, pdf_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully",
                                                font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                                fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                def download_file():
                    save_path = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                             filetypes=[("PDF Files", "*.pdf")])

                    if save_path:
                        os.rename(pdf_path, save_path)
                        messagebox.showinfo("Download Complete", "PDF (.pdf) file downloaded successfully.")
                    else:
                        messagebox.showerror("Save Error", "Unsupported file format.")

                # Download Button
                download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                            pady=5, bg=right_frame["bg"], fg="white",
                                            activebackground="green", highlightcolor="green",
                                            command=download_file)
                download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Adding upload button
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 2 : CSV to Excel


def show_csv_to_excel():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your Csv (.csv) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("Csv Files", "*.csv")])
        if file_path:
            if not file_path.lower().endswith('.csv'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the csv file to excel format
                excel_path = os.path.splitext(file_path)[0] + ".xlsx"
                success = convert_csv_to_excel(file_path, excel_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully",
                                                font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                                fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                def download_file():
                    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                             filetypes=[("Excel Files", "*.xlsx")])

                    if save_path:
                        os.rename(excel_path, save_path)
                        messagebox.showinfo("Download Complete", "Excel (.xlsx) file downloaded successfully.")
                    else:
                        messagebox.showerror("Save Error", "Unsupported file format.")

                # Download Button
                download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                            pady=5, bg=right_frame["bg"], fg="white",
                                            activebackground="green", highlightcolor="green",
                                            command=download_file)
                download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Adding upload button
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Button 3 : Excel to CSV


def show_excel_to_csv():
    # Configure the right frame background color
    right_frame.configure(bg=app.cget("background"))

    # Add the content for PDF to PNG conversion
    box_frame = tk.Frame(right_frame, width=1200, height=900, bg=app.cget("background"),
                         highlightbackground="SpringGreen4", highlightthickness=3)
    box_frame.place(relx=0.5, rely=0.3, anchor=tk.CENTER)

    # Add the text inside the box
    text_label = tk.Label(box_frame, text="Upload your Excel (.xlsx) file here.....", font=("Arial", 16),
                          padx=20, pady=20, bg=app.cget("background"), fg="gold2")
    text_label.pack()

    # Function to handle file upload
    def upload_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            if not file_path.lower().endswith('.xlsx'):
                messagebox.showerror("Error", "Selected file is not in a supported format.")
            else:
                # Clear any existing file name label and buttons
                if 'filename_label' in locals():
                    filename_label.destroy()
                if 'conversion_label' in locals():
                    conversion_label.destroy()
                if 'download_button' in locals():
                    download_button.destroy()

                # Display the file name below the rectangular box
                filename_label = tk.Label(right_frame, text="File: " + os.path.basename(file_path),
                                          font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                          fg="white")
                filename_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

                # Convert the csv file to excel format
                csv_path = os.path.splitext(file_path)[0] + ".xlsx"
                success = convert_excel_to_csv(file_path, csv_path)

                if success:
                    conversion_label = tk.Label(right_frame, text="File converted successfully",
                                                font=("Arial", 14), padx=20, pady=10, bg=app.cget("background"),
                                                fg="green")
                    conversion_label.place(relx=0.5, rely=0.6, anchor=tk.CENTER)

                def download_file():
                    save_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                                             filetypes=[("CSV Files", "*.csv")])

                    if save_path:
                        os.rename(csv_path, save_path)
                        messagebox.showinfo("Download Complete", "Excel (.xlsx) file downloaded successfully.")
                    else:
                        messagebox.showerror("Save Error", "Unsupported file format.")

                # Download Button
                download_button = tk.Button(right_frame, text="Download", font=("Arial", 14), padx=10,
                                            pady=5, bg=right_frame["bg"], fg="white",
                                            activebackground="green", highlightcolor="green",
                                            command=download_file)
                download_button.place(relx=0.5, rely=0.7, anchor=tk.CENTER)

    # Adding upload button
    upload_button = ctk.CTkButton(box_frame, text="Upload File", image=upload_photo, font=ctk.CTkFont(size=14),
                                  compound=tk.LEFT, fg_color="transparent", command=upload_file)
    upload_button.configure(text_color="white")
    upload_button.pack(pady=10)


# Set the appearance mode and default color theme
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("green")


# Our app frame
app = ctk.CTk()
app.geometry("1100x580")
app.title("Easy Convert Tool")

# App icon
app_icon = Path(__file__).resolve().parent / 'title_img.ico'

app.iconbitmap(app_icon)


# Left Scrollable frame
scrollable_frame = ctk.CTkScrollableFrame(app, width=250, scrollbar_button_color="white", fg_color="black",
                                          scrollbar_button_hover_color="green", corner_radius=0)
scrollable_frame.pack(side="left", fill="y")

# Create a PhotoImage object for each conversion button
b_path = Path(__file__).resolve().parent / 'icons8-arrow-16.png'
button_img = Image.open(b_path)
button_img = button_img.resize((18, 19))  # Resize the image if needed
button_photo = ImageTk.PhotoImage(button_img)

# Create a UploadImage object for each upload image
u_path = Path(__file__).resolve().parent / 'upload_button.png'
upload_img = Image.open(u_path)
upload_img = upload_img.resize((16, 16))  # Resize the image if needed
upload_photo = ImageTk.PhotoImage(upload_img)

# Load home icon image
h_path = Path(__file__).resolve().parent / 'desktop_icon.png'
home_image = Image.open(h_path)  # Replace "home_icon.png" with your
# image file path
home_image = home_image.resize((34, 34), Image.LANCZOS)
home_icon = ImageTk.PhotoImage(home_image)

# Home Button
home_button = tk.Button(scrollable_frame, image=home_icon, bd=0, command=go_to_home)
home_button.config(highlightthickness=0, highlightbackground=app.cget("background"))
home_button.config(width=30, height=30)
home_button.pack(anchor="w", padx=15, pady=(5, 10))

# Labels for left Frame Panel
logo_label = ctk.CTkLabel(scrollable_frame, text="Image Conversions", fg_color="transparent",
                          font=ctk.CTkFont(size=16, weight="bold", underline=True))
logo_label.pack(anchor="w", padx=15)

# Button Frame
button_frame = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
button_frame.pack(anchor="w", pady=15)

# Button 1: PNG to JPEG
button1 = ctk.CTkButton(button_frame, text="PNG to JPEG", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_png_to_jpeg)
button1.pack(anchor="w", padx=15)

# Button 2 : JPEG to PNG
button2 = ctk.CTkButton(button_frame, text="JPEG to PNG", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_jpeg_to_png)
button2.pack(anchor="w", padx=15)

# Button 3 : PNG\JPEG to PDF
button3 = ctk.CTkButton(button_frame, text="PNG\JPEG TO PDF", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_image_to_pdf)
button3.pack(anchor="w", padx=25)

# Button 4 : PDF to JPEG or PNG format
button4 = ctk.CTkButton(button_frame, text="PDF TO PNG\JPEG", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_pdf_to_png)
button4.pack(anchor="w", padx=25)

# Section 2 : Document Conversion
# Second label for the next section
frame2 = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
frame2.pack(anchor="w", pady=20)

# Labels for the second section
logo_label1 = ctk.CTkLabel(frame2, text="Document Conversions", fg_color="transparent",
                           font=ctk.CTkFont(size=16, weight="bold", underline=True))
logo_label1.pack(anchor="w", padx=10, pady=10)

# Button 1: PDF TO WORD
button1 = ctk.CTkButton(frame2, text="PDF to Word", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_pdf_to_word)

button1.pack(anchor="w", padx=13)

# Button 2: WORD TO PDF Format
button2 = ctk.CTkButton(frame2, text="Word to PDF", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_word_to_pdf)

button2.pack(anchor="w", padx=12)

# Button 3: EXCEL TO PDF Format
button3 = ctk.CTkButton(frame2, text="EXCEL to PDF", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_excel_to_pdf)

button3.pack(anchor="w", padx=14)

# Button 4: EXCEL TO PDF Format
button4 = ctk.CTkButton(frame2, text="PDF to EXCEL", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_pdf_to_excel)

button4.pack(anchor="w", padx=14)

# Button 5: PDF TO PowerPoint Format
button5 = ctk.CTkButton(frame2, text="PDF to PowerPoint", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_pdf_to_ppt)

button5.pack(anchor="w", padx=14)

# Button 6: PowerPoint TO PDF Format
button6 = ctk.CTkButton(frame2, text="PowerPoint to PDF", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_ppt_to_pdf)

button6.pack(anchor="w", padx=14)

# Section 2
# File Format Conversions

# Third label for the next section
frame3 = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
frame3.pack(anchor="w", pady=20)

# Labels for the third section
logo_label2 = ctk.CTkLabel(frame3, text="File Format Conversions", fg_color="transparent",
                           font=ctk.CTkFont(size=16, weight="bold", underline=True))
logo_label2.pack(anchor="w", padx=10)

# Button 1 : Text to PDF format

button1 = ctk.CTkButton(frame3, text="Text to PDF", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_text_to_pdf)
button1.pack(anchor="w", padx=10)

# Button 2 : CSV to Excel format

button1 = ctk.CTkButton(frame3, text="CSV to Excel", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_csv_to_excel)
button1.pack(anchor="w", padx=10)

# Button 3 : Excel to CSV format

button1 = ctk.CTkButton(frame3, text="Excel to CSV", image=button_photo, fg_color="transparent",
                        font=ctk.CTkFont(size=14), compound=tk.LEFT, command=show_excel_to_csv)
button1.pack(anchor="w", padx=10)

# Right Frame
right_frame = tk.Frame(app, width=850, bg=app.cget("background"), highlightbackground=app.cget("background"),
                       highlightthickness=3)
right_frame.pack(side="right", fill="both", expand=True)

# Opening Right Frame Text
title_name = "Easy Convert Tool"
powered_by = "CortexX"
disclaimer = "This tool is for personal use only."
contact_details = "For any queries, contact me on:"
github_account = "https://github.com/CodeWithMayank-Py"
gmail_account = "paliwalm4321@gmail.com"
insta_account = "@mayank._.paliwal"

# Create a label widget for the title
title_label = tk.Label(right_frame, text=title_name, font=("Courier New", 54, "bold italic underline"),
                       bg=app.cget("background"), fg="white")
title_label.pack(fill=tk.BOTH, expand=True)

# Create a label widget for the "Powered by" text
powered_by_label = tk.Label(right_frame, text=f"Powered by {powered_by}", font=("Courier New", 24, "italic"),
                            bg=app.cget("background"), fg="green2")
powered_by_label.pack()

# Create a label widget for the disclaimer
disclaimer_label = tk.Label(right_frame, text=f"Disclaimer: {disclaimer}", font=("Trebuchet MS", 14, "italic"),
                            bg=app.cget("background"), fg="red3")
disclaimer_label.pack()

# GitHub Label and Logo
# Load the images for GitHub
g_path = Path(__file__).resolve().parent / 'github_icon.png'
git_image = Image.open(g_path)
# image file path
git_image = git_image.resize((32, 32), Image.LANCZOS)
github_image = ImageTk.PhotoImage(git_image)

# GitHub Label
github_label = tk.Label(right_frame, image=github_image, text=f"{github_account}", font=("Futura", 10),
                        bg=app.cget("background"),
                        fg="white", compound=tk.LEFT, padx=5)
github_label.pack(side="left", anchor=tk.E, padx=160)

# Gmail Logo and Label
# Load the Gmail Image
gm_path = Path(__file__).resolve().parent / 'gmail_icon.png'
gm_image = Image.open(gm_path)
# Image file path
gm_image = gm_image.resize((28, 28), Image.LANCZOS)
gmail_image = ImageTk.PhotoImage(gm_image)

gmail_label = tk.Label(right_frame, image=gmail_image, text=f"{gmail_account}", font=("Futura", 10),
                       bg=app.cget("background"),
                       fg="white", compound=tk.LEFT, padx=5)
gmail_label.pack(side="left", anchor=tk.E) 

# For music
music_path = Path(__file__).resolve().parent / 'too-late-now.wav'
pygame.mixer.init()
pygame.mixer.music.load(music_path)
pygame.mixer.music.set_volume(0.4)
pygame.mixer.music.play(-1)


# Code to stop music and quit pygame when the app closed

def close_app():
    pygame.mixer.music.stop()
    pygame.quit()
    app.destroy()


app.protocol("WM_DELETE_WINDOW", close_app)

# ================================================================================================================= #

# ================================================================================================================== #

app.mainloop()
