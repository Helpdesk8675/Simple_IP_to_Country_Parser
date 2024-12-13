#  Need to install the following
#  pip install pdfminer.six
#  pip install python-docx
#  pip install olefile

import tkinter as tk
from tkinter import filedialog, messagebox
import csv
import re
import os
from datetime import datetime
import ipaddress

try:
    from pdfminer.high_level import extract_text
except ImportError:
    extract_text = None
    print("Warning: pdfminer.six is not installed. PDF processing will not be available.")

try:
    from docx import Document
except ImportError:
    Document = None
    print("Warning: python-docx is not installed. Word document processing will not be available.")

try:
    import olefile
except ImportError:
    olefile = None
    print("Warning: olefile is not installed. Binary file processing will not be available.")

class IPSearchGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("IP Address Search Tool")
        self.root.geometry("600x400")

        # Variables to store paths
        self.folder_path = tk.StringVar()
        self.ipv4_ref_path = tk.StringVar()
        self.ipv6_ref_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.create_gui()

    def create_gui(self):
        # Input folder selection
        input_frame = tk.Frame(self.root)
        input_frame.pack(pady=5, fill="x")
        tk.Label(input_frame, text="Select Input Folder:").pack(side="left", padx=5)
        tk.Entry(input_frame, textvariable=self.folder_path, width=50).pack(side="left", padx=5)
        tk.Button(input_frame, text="Browse", command=self.browse_folder).pack(side="left", padx=5)

        # IPv4 reference file selection
        ipv4_frame = tk.Frame(self.root)
        ipv4_frame.pack(pady=5, fill="x")
        tk.Label(ipv4_frame, text="Select IPv4 Reference File:").pack(side="left", padx=5)
        tk.Entry(ipv4_frame, textvariable=self.ipv4_ref_path, width=50).pack(side="left", padx=5)
        tk.Button(ipv4_frame, text="Browse", command=self.browse_ipv4).pack(side="left", padx=5)

        # IPv6 reference file selection
        ipv6_frame = tk.Frame(self.root)
        ipv6_frame.pack(pady=5, fill="x")
        tk.Label(ipv6_frame, text="Select IPv6 Reference File:").pack(side="left", padx=5)
        tk.Entry(ipv6_frame, textvariable=self.ipv6_ref_path, width=50).pack(side="left", padx=5)
        tk.Button(ipv6_frame, text="Browse", command=self.browse_ipv6).pack(side="left", padx=5)

        # Output folder selection
        output_frame = tk.Frame(self.root)
        output_frame.pack(pady=5, fill="x")
        tk.Label(output_frame, text="Select Output Folder:").pack(side="left", padx=5)
        tk.Entry(output_frame, textvariable=self.output_path, width=50).pack(side="left", padx=5)
        tk.Button(output_frame, text="Browse", command=self.browse_output).pack(side="left", padx=5)

        # Process button
        tk.Button(self.root, text="Process Files", command=self.process_files).pack(pady=20)

    def browse_folder(self):
        folder = filedialog.askdirectory()
        self.folder_path.set(folder)

    def browse_ipv4(self):
        file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.ipv4_ref_path.set(file)

    def browse_ipv6(self):
        file = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.ipv6_ref_path.set(file)

    def browse_output(self):
        folder = filedialog.askdirectory()
        self.output_path.set(folder)

    def load_ip_references(self):
        ip_ranges = []
        try:
            # Load IPv4 references
            with open(self.ipv4_ref_path.get(), 'r') as f:
                reader = csv.reader(f)
                headers = next(reader)  # Skip header
                if len(headers) < 3:
                    raise ValueError("IPv4 file must have three columns: 'Start IP', 'End IP', and 'Country Code'.")
                for row in reader:
                    try:
                        start_ip = ipaddress.ip_address(row[0])  # Start of range
                        end_ip = ipaddress.ip_address(row[1])    # End of range
                        ip_ranges.append((start_ip, end_ip, row[2]))  # Append tuple (start_ip, end_ip, country)
                    except ValueError:
                        continue

            # Load IPv6 references
            with open(self.ipv6_ref_path.get(), 'r') as f:
                reader = csv.reader(f)
                headers = next(reader)  # Skip header
                if len(headers) < 3:
                    raise ValueError("IPv6 file must have three columns: 'Start IP', 'End IP', and 'Country Code'.")
                for row in reader:
                    try:
                        start_ip = ipaddress.ip_address(row[0])
                        end_ip = ipaddress.ip_address(row[1])
                        ip_ranges.append((start_ip, end_ip, row[2]))
                    except ValueError:
                        continue

        except FileNotFoundError as e:
            messagebox.showerror("Error", f"Reference file not found: {e}")
        except ValueError as e:
            messagebox.showerror("Error", f"Invalid reference file format: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error loading references: {e}")

        return ip_ranges

    def find_ips(self, content):
        # Extract IPv4 and IPv6 addresses from the content using regex
        ipv4_pattern = r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b'
        ipv6_pattern = r'\b(?:[A-Fa-f0-9]{1,4}:){1,7}[A-Fa-f0-9]{1,4}\b'
        ipv4_matches = re.findall(ipv4_pattern, content)
        ipv6_matches = re.findall(ipv6_pattern, content)
        return ipv4_matches + ipv6_matches

    def get_country_for_ip(self, ip_address, ip_ranges):
        try:
            ip_obj = ipaddress.ip_address(ip_address)
            for start_ip, end_ip, country in ip_ranges:
                if start_ip <= ip_obj <= end_ip:  # Check if IP is in range
                    return country
            return "Unknown"
        except ValueError:
            return "Invalid IP"

    def extract_content_from_file(self, file_path):
        content = ""
        try:
            if file_path.lower().endswith('.pdf'):
                if extract_text:
                    content = extract_text(file_path)
                else:
                    print("PDF processing is unavailable. Skipping file.")
            elif file_path.lower().endswith('.docx'):
                if Document:
                    doc = Document(file_path)
                    content = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                else:
                    print("Word document processing is unavailable. Skipping file.")
            elif file_path.lower().endswith('.xls') or file_path.lower().endswith('.xlsx'):
                if olefile and olefile.isOleFile(file_path):
                    with olefile.OleFileIO(file_path) as ole:
                        for stream in ole.listdir():
                            content += str(ole.openstream(stream).read())
                else:
                    print("Binary file processing is unavailable or unsupported for this file. Skipping file.")
            else:
                with open(file_path, 'r', errors='ignore', encoding='utf-8') as current_file:
                    content = current_file.read()
        except Exception as e:
            print(f"Error reading file {file_path}: {e}")
        return content

    def process_files(self):
        try:
            # Validate inputs
            if not all([self.folder_path.get(), self.ipv4_ref_path.get(), 
                        self.ipv6_ref_path.get(), self.output_path.get()]):
                messagebox.showerror("Error", "All fields are required!")
                return

            # Load IP reference data
            ip_references = self.load_ip_references()
            if not ip_references:
                return

            # Create output filename
            timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
            output_file = os.path.join(self.output_path.get(), f"{timestamp}_ipsearch.csv")

            # Process files and write results
            with open(output_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(['Filename', 'IP Address', 'Country'])

                # Recursively search through files
                for root, _, files in os.walk(self.folder_path.get()):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            content = self.extract_content_from_file(file_path)
                            ips = self.find_ips(content)
                            for ip in ips:
                                country = self.get_country_for_ip(ip, ip_references)
                                writer.writerow([file, ip, country])
                        except Exception as e:
                            print(f"Error processing file {file}: {str(e)}")
                            continue

            messagebox.showinfo("Success", f"Processing complete! Results saved to {output_file}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Main entry point
def main():
    root = tk.Tk()
    app = IPSearchGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
