import os
import difflib
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

class DuplicateFinderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Duplicate File Finder")
        self.source_folder = None
        self.destination_file = None
        self.total_files = 0
        self.processed_files = 0
        
        self.create_widgets()
        
    def create_widgets(self):
        # Source folder selection
        tk.Button(self.root, text="Source", command=self.browse_source).grid(row=0, column=0, padx=10, pady=10)
        self.source_label = tk.Label(self.root, text="Source: Not selected", anchor="w")
        self.source_label.grid(row=0, column=1, padx=10, sticky="ew")
        
        # Destination file selection
        tk.Button(self.root, text="Destination", command=self.browse_destination).grid(row=1, column=0, padx=10, pady=10)
        self.destination_label = tk.Label(self.root, text="Destination: Not selected", anchor="w")
        self.destination_label.grid(row=1, column=1, padx=10, sticky="ew")
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        
        # Progress label
        self.progress_label = tk.Label(self.root, text="")
        self.progress_label.grid(row=3, column=0, columnspan=2, padx=10)
        
        # Execute button
        tk.Button(self.root, text="OK", command=self.execute_script).grid(row=4, column=0, columnspan=2, pady=20)
        
        # Result label
        self.result_label = tk.Label(self.root, text="", anchor="w")
        self.result_label.grid(row=5, column=0, columnspan=2, padx=10)

    def browse_source(self):
        folder = filedialog.askdirectory(title="Select Source Folder")
        if folder:
            self.source_folder = folder
            self.source_label.config(text=f"Source: {folder}")

    def browse_destination(self):
        file = filedialog.asksaveasfilename(
            title="Save Excel File",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file:
            self.destination_file = file
            self.destination_label.config(text=f"Destination: {file}")

    def find_similar_files(self):
        files = {}
        for root, _, filenames in os.walk(self.source_folder):
            for filename in filenames:
                filepath = os.path.join(root, filename)
                try:
                    size = os.path.getsize(filepath)
                    files.setdefault(size, []).append(filepath)
                    self.total_files += 1
                except PermissionError:
                    continue
        
        # Update progress bar
        self.progress_bar['maximum'] = self.total_files
        self.progress_bar['value'] = 0
        self.progress_label.config(text="Processing files...")
        
        # Find size duplicates
        size_duplicates = {}
        for size, paths in files.items():
            if len(paths) > 1:
                size_duplicates[size] = paths
                for path in paths:
                    self.processed_files += 1
                    self.progress_bar['value'] = self.processed_files
                    self.progress_label.config(text=f"Processing {self.processed_files}/{self.total_files} files")
                    self.root.update_idletasks()
        
        # Find filename similarities
        filename_similarities = {}
        for paths in size_duplicates.values():
            for path in paths:
                filename = os.path.basename(path)
                for other_path in paths:
                    if path != other_path:
                        other_filename = os.path.basename(other_path)
                        match = difflib.SequenceMatcher(None, filename, other_filename).find_longest_match(
                            0, len(filename), 0, len(other_filename)
                        )
                        if match.size >= 5:
                            key = tuple(sorted([path, other_path]))
                            filename_similarities.setdefault(key, []).append(match)
        
        return size_duplicates, filename_similarities

    def save_to_excel(self, size_duplicates, filename_similarities, output_file):
        # Prepare data for Excel
        size_duplicates_data = []
        for size, paths in size_duplicates.items():
            size_duplicates_data.append({"Size (bytes)": size, "Files": ", ".join(paths)})

        filename_similarities_data = []
        for (path1, path2), matches in filename_similarities.items():
            match_details = "; ".join([f"{match.a}:{match.a+match.size}" for match in matches])
            filename_similarities_data.append({"File 1": path1, "File 2": path2, "Match Details": match_details})

        # Create DataFrames
        df_size_duplicates = pd.DataFrame(size_duplicates_data)
        df_filename_similarities = pd.DataFrame(filename_similarities_data)

        # Save to Excel
        with pd.ExcelWriter(output_file) as writer:
            df_size_duplicates.to_excel(writer, sheet_name="Size Duplicates", index=False)
            df_filename_similarities.to_excel(writer, sheet_name="Filename Similarities", index=False)

    def execute_script(self):
        if not self.source_folder or not self.destination_file:
            messagebox.showerror("Error", "Please select both source and destination!")
            return

        try:
            size_duplicates, filename_similarities = self.find_similar_files()
            self.save_to_excel(size_duplicates, filename_similarities, self.destination_file)
            self.result_label.config(text=f"Results saved to {self.destination_file}", fg="green")
        except Exception as e:
            self.result_label.config(text=f"Error: {str(e)}", fg="red")
        finally:
            self.progress_bar['value'] = 0
            self.progress_label.config(text="")
            self.total_files = 0
            self.processed_files = 0

if __name__ == "__main__":
    root = tk.Tk()
    app = DuplicateFinderGUI(root)
    root.mainloop()

