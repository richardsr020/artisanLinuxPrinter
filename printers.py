import os
import cups
import fitz  # PyMuPDF
import tkinter as tk
from tkinter import ttk, messagebox


class LinuxPrinterService:
    def __init__(self, root, pdf_path, password):
        self.root = root
        self.root.title("Secure PDF Printing")
        self.root.geometry("500x500")

        self.pdf_path = pdf_path
        self.password = password
        self.unlocked_pdf = None  # Stores the unlocked PDF file
        self.total_pages = 0  # Total number of pages in the PDF

        self.unlock_pdf()

        # Connect to CUPS and get available printers
        self.conn = cups.Connection()
        self.printers = list(self.conn.getPrinters().keys())

        if not self.printers:
            messagebox.showerror("Error", "No printers detected.")
            return

        # Creating the window
        frame_main = tk.Frame(root)
        frame_main.pack(pady=10)

        # Printer selection
        frame_printer = tk.LabelFrame(frame_main, text="Printer", padx=10, pady=10)
        frame_printer.grid(row=0, column=0, padx=10, pady=5)

        tk.Label(frame_printer, text="Select a printer:").grid(row=0, column=0)
        self.printer_var = tk.StringVar(root, self.printers[0])
        self.dropdown = ttk.Combobox(frame_printer, textvariable=self.printer_var, values=self.printers, state="readonly")
        self.dropdown.grid(row=1, column=0, pady=5)

        # Number of copies
        frame_copies = tk.LabelFrame(frame_main, text="Number of Copies", padx=10, pady=10)
        frame_copies.grid(row=1, column=0, padx=10, pady=5)

        self.copies_var = tk.IntVar(value=1)
        self.copies_entry = tk.Entry(frame_copies, textvariable=self.copies_var)
        self.copies_entry.grid(row=0, column=0, pady=5)

        # Select pages to print
        frame_pages = tk.LabelFrame(frame_main, text="Pages to Print", padx=10, pady=10)
        frame_pages.grid(row=2, column=0, padx=10, pady=5)

        tk.Label(frame_pages, text="Pages (e.g., 1-5,7,10-15):").grid(row=0, column=0)
        self.pages_var = tk.StringVar()
        self.pages_entry = tk.Entry(frame_pages, textvariable=self.pages_var)
        self.pages_entry.grid(row=1, column=0, pady=5)

        # Set the default value to match total pages of the document
        self.pages_var.set(f"1-{self.total_pages}")

        # Even/Odd page mode
        tk.Label(frame_pages, text="Page Mode:").grid(row=2, column=0)
        self.page_mode_var = tk.StringVar(value="all")
        self.page_mode_dropdown = ttk.Combobox(frame_pages, textvariable=self.page_mode_var, values=["All", "Even", "Odd"], state="readonly")
        self.page_mode_dropdown.grid(row=3, column=0, pady=5)

        # Additional printing options
        frame_options = tk.LabelFrame(frame_main, text="Printing Options", padx=10, pady=10)
        frame_options.grid(row=3, column=0, padx=10, pady=5)

        # Paper orientation
        tk.Label(frame_options, text="Orientation:").grid(row=0, column=0)
        self.orientation_var = tk.StringVar(value="portrait")
        self.orientation_dropdown = ttk.Combobox(frame_options, textvariable=self.orientation_var, values=["Portrait", "Landscape"], state="readonly")
        self.orientation_dropdown.grid(row=0, column=1, pady=5)

        # Paper size
        tk.Label(frame_options, text="Paper Size:").grid(row=1, column=0)
        self.size_var = tk.StringVar(value="A4")
        self.size_dropdown = ttk.Combobox(frame_options, textvariable=self.size_var, values=["A4", "A3", "Letter", "Legal"], state="readonly")
        self.size_dropdown.grid(row=1, column=1, pady=5)

        # Print button
        self.btn_print = tk.Button(frame_main, text="Print", command=self.print_pdf)
        self.btn_print.grid(row=4, column=0, pady=20)

    def unlock_pdf(self):
        """Unlock the PDF file with the password and retrieve the total number of pages."""
        try:
            doc = fitz.open(self.pdf_path)

            if doc.needs_pass and not doc.authenticate(self.password):
                messagebox.showerror("Error", "Incorrect password!")
                self.root.destroy()
                return

            self.total_pages = len(doc)

            # Create a temporary unlocked PDF file
            self.unlocked_pdf = self.pdf_path.replace(".pdf", "_unlocked.pdf")
            doc.save(self.unlocked_pdf)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to unlock the PDF: {e}")
            self.root.destroy()

    def parse_page_selection(self, selection):
        """Parse and validate the user's input for pages to print."""
        selected_pages = set()

        if not selection.strip():
            return set(range(1, self.total_pages + 1))  # Print all pages if no filter

        parts = selection.split(',')
        for part in parts:
            if '-' in part:  # Handle ranges like "1-5"
                try:
                    start, end = map(int, part.split('-'))
                    if start > end or start < 1 or end > self.total_pages:
                        raise ValueError
                    selected_pages.update(range(start, end + 1))
                except ValueError:
                    messagebox.showerror("Error", f"Invalid page range: {part}")
                    return None
            else:  # Handle individual pages like "3,7,10"
                try:
                    page = int(part)
                    if page < 1 or page > self.total_pages:
                        raise ValueError
                    selected_pages.add(page)
                except ValueError:
                    messagebox.showerror("Error", f"Invalid page: {part}")
                    return None

        return selected_pages

    def print_pdf(self):
        """Print the unlocked PDF file with the user's selected options."""
        if not self.unlocked_pdf:
            messagebox.showerror("Error", "The PDF file was not unlocked.")
            return

        try:
            copies = int(self.copies_var.get())
            if copies < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Error", "Number of copies must be a positive integer.")
            return

        selected_pages = self.parse_page_selection(self.pages_var.get())
        if selected_pages is None:
            return

        # Filter by even/odd pages
        mode = self.page_mode_var.get()
        if mode == "Even":
            selected_pages = {p for p in selected_pages if p % 2 == 0}
        elif mode == "Odd":
            selected_pages = {p for p in selected_pages if p % 2 == 1}

        if not selected_pages:
            messagebox.showerror("Error", "No valid pages to print.")
            return

        # Convert page numbers to a format suitable for CUPS
        page_list = ",".join(map(str, sorted(selected_pages)))
        selected_printer = self.printer_var.get()

        # Printing options
        print_options = {
            "copies": str(copies),
            "page-ranges": page_list,
            "orientation-requested": "3" if self.orientation_var.get() == "portrait" else "4",
            "media": self.size_var.get().lower()
        }

        # Send the print job
        self.conn.printFile(selected_printer, self.unlocked_pdf, "PDF Print", print_options)
        messagebox.showinfo("Success", f"The document has been sent to the printer.\nPages: {page_list}")

        # Delete the temporary unlocked PDF file after printing
        try:
            os.remove(self.unlocked_pdf)
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not delete the temporary file: {e}")


# # Exécution de l'application Tkinter avec des arguments
# if __name__ == "__main__":
#     root = tk.Tk()

#     root.mainloop()
