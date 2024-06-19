import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Entry, Button, Text, ttk
import pyexcel as p

class EligibilityApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Eligibility List Processor")

        # File paths
        self.eligibility_file = None
        self.exclusion_file = None

        # Data
        self.eligibility_df = None
        self.exclusion_df = None
        self.filtered_df = None
        self.subject_columns = None
        self.weights = []
        self.max_marks = []

        # GUI elements
        self.create_widgets()

    def create_widgets(self):
        # Upload Eligibility List Button
        self.upload_eligibility_button = Button(self.root, text="Upload Student Marks Excel", command=self.upload_eligibility)
        self.upload_eligibility_button.grid(row=1, column=0, padx=10, pady=10)

        # Upload Exclusion List Button
        self.upload_exclusion_button = Button(self.root, text="Upload Exclusion Excel", command=self.upload_exclusion)
        self.upload_exclusion_button.grid(row=1, column=1, padx=10, pady=10)

        # Process Button
        self.process_button = Button(self.root, text="Process", command=self.process_data)
        self.process_button.grid(row=1, column=2, padx=10, pady=10)
        self.process_button.config(state="disabled")

        # Reset Button
        self.reset_button = Button(self.root, text="Reset", command=self.reset_app)
        self.reset_button.grid(row=1, column=3, padx=10, pady=10)

        # Guideline Section
        self.guideline_text = Text(self.root, height=15, width=80, wrap=tk.WORD, bg=self.root.cget('bg'), bd=0)
        self.guideline_text.insert(tk.END, 
            "Guidelines:\n"
            "1. Click 'Upload Student Marks Excel' to upload the student marks file. The file should have columns 'Index Number', 'Student Name', and subject marks.\n"
            "2. Click 'Upload Exclusion Excel' to upload the exclusion list file. The file should have a column 'Index Number'.\n"
            "3. After uploading both files, click 'Process'.\n"
            "4. Enter the weights and maximum marks for each subject, then click 'Submit'.\n"
            "5. The application will calculate the final marks and determine eligibility.\n"
            "6. Save the final eligibility list to your desired location."
        )
        self.guideline_text.config(state=tk.DISABLED)
        self.guideline_text.grid(row=0, column=0, columnspan=4, padx=10, pady=10)

    def upload_file(self, file_type):
        file_path = filedialog.askopenfilename(title=f"Upload {file_type} Excel File", filetypes=[("Excel files", "*.xlsx *.xls *.ods")])
        if not file_path:
            messagebox.showerror("Error", "No file selected.")
            return None
        return file_path

    def upload_eligibility(self):
        self.eligibility_file = self.upload_file("Eligibility List")
        if self.eligibility_file:
            self.eligibility_df = self.read_file(self.eligibility_file)
            if self.eligibility_df is not None:
                self.upload_eligibility_button.config(text="Eligibility List Uploaded", state="disabled")
                messagebox.showinfo("Success", "Student mark list uploaded successfully.")
                self.check_files_uploaded()

    def upload_exclusion(self):
        self.exclusion_file = self.upload_file("Exclusion List")
        if self.exclusion_file:
            self.exclusion_df = self.read_file(self.exclusion_file)
            if self.exclusion_df is not None:
                self.upload_exclusion_button.config(text="Exclusion List Uploaded", state="disabled")
                messagebox.showinfo("Success", "Exclusion list uploaded successfully.")
                self.check_files_uploaded()

    def read_file(self, file_path):
        try:
            if file_path.endswith('.xls') or file_path.endswith('.xlsx'):
                return pd.read_excel(file_path)
            elif file_path.endswith('.ods'):
                return pd.DataFrame(p.get_sheet(file_name=file_path).to_records())
            else:
                messagebox.showerror("Error", "Unsupported file format.")
                return None
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read file: {e}")
            return None

    def check_files_uploaded(self):
        if self.eligibility_file and self.exclusion_file:
            self.process_button.config(state="normal")

    def get_subject_weights_and_max_marks(self):
        weight_entries = []
        max_mark_entries = []
        for i, column in enumerate(self.subject_columns):
            Label(self.root, text=f"Weight for {column}:").grid(row=i + 2, column=0, padx=10, pady=5)
            weight_entry = Entry(self.root)
            weight_entry.grid(row=i + 2, column=1, padx=10, pady=5)
            weight_entries.append(weight_entry)

            Label(self.root, text=f"Max marks for {column}:").grid(row=i + 2, column=2, padx=10, pady=5)
            max_mark_entry = Entry(self.root)
            max_mark_entry.grid(row=i + 2, column=3, padx=10, pady=5)
            max_mark_entries.append(max_mark_entry)

        def submit_weights_and_marks():
            try:
                self.weights = [float(entry.get()) for entry in weight_entries]
                self.max_marks = [float(entry.get()) for entry in max_mark_entries]
                self.calculate_final_marks()
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numerical values for weights and max marks.")
                return

        Button(self.root, text="Submit", command=submit_weights_and_marks).grid(row=len(self.subject_columns) + 2, column=1, padx=10, pady=10)

    def process_data(self):
        excluded_indices = self.exclusion_df["Index Number"].tolist()
        self.filtered_df = self.eligibility_df[~self.eligibility_df["Index Number"].isin(excluded_indices)]
        self.subject_columns = self.filtered_df.columns[2:]  # Assuming first two columns are Index Number and Student Name
        self.get_subject_weights_and_max_marks()

    def calculate_final_marks(self):
        # Convert non-float values to 0 in subject columns
        for column in self.subject_columns:
            self.filtered_df[column] = pd.to_numeric(self.filtered_df[column], errors='coerce').fillna(0)

        def calculate_final(row):
            total_weight = sum(self.weights)
            weighted_sum = sum((row[subject] / max_mark * weight) for subject, max_mark, weight in zip(self.subject_columns, self.max_marks, self.weights))
            final_marks = (weighted_sum / total_weight) * 100
            return final_marks

        self.filtered_df['Grade'] = self.filtered_df.apply(calculate_final, axis=1)
        self.filtered_df['Eligibility'] = self.filtered_df['Grade'].apply(lambda x: "Eligible" if x >= 40 else "Not Eligible")

        # Sort the DataFrame by "Index Number" in ascending order
        self.filtered_df.sort_values(by="Index Number", inplace=True)

        # Display the data in a table
        self.display_data()

        # Enable the download button
        self.download_button = Button(self.root, text="Download Final Document", command=self.save_final_document)
        self.download_button.grid(row=len(self.subject_columns) + 3, column=1, padx=10, pady=10)

    def display_data(self):
        cols = list(self.filtered_df.columns)
        tree = ttk.Treeview(self.root, columns=cols, show='headings')
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, minwidth=0, width=100)
        for index, row in self.filtered_df.iterrows():
            tree.insert("", "end", values=list(row))
        tree.grid(row=len(self.subject_columns) + 4, column=0, columnspan=4, padx=10, pady=10)

    def save_final_document(self):
        # Prompt user to select a location and name for the output file
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx *.xls")], title="Save Final Eligibility List")
        if output_file:
            self.filtered_df.to_excel(output_file, index=False)
            messagebox.showinfo("Success", f"Final eligibility list created successfully!\nSaved as {output_file}")
        else:
            messagebox.showwarning("Warning", "No file was selected. The final eligibility list was not saved.")

    def reset_app(self):
        self.eligibility_file = None
        self.exclusion_file = None
        self.eligibility_df = None
        self.exclusion_df = None
        self.filtered_df = None
        self.subject_columns = None
        self.weights = []
        self.max_marks = []

        for widget in self.root.winfo_children():
            widget.destroy()

        self.create_widgets()

if __name__ == "__main__":
    root = tk.Tk()
    app = EligibilityApp(root)
    root.mainloop()
