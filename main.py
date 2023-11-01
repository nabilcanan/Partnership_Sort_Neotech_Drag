import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from compare_neotech_drop import compare_neotech_drag


class DropApp(tk.Frame):
    def __init__(self, root):
        super().__init__(root)

        # Design elements
        self.pack(expand=True, fill="both")
        label = tk.Label(self, text="Drop the new week's file folder here", padx=10, pady=10)
        label.pack(pady=150, padx=150)

        # Bind drag and drop event
        root.drop_target_register(DND_FILES)
        root.dnd_bind('<<Drop>>', self.drop)

    def drop(self, event):
        folder_path = event.data
        print("Folder dropped:", folder_path)

        if os.path.isdir(folder_path):
            xls_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if
                         file.lower().endswith('.xls')]
            xls_files = sorted(xls_files, key=lambda x: os.path.basename(x), reverse=True)

            # Ensuring there are at least two files
            if len(xls_files) >= 2:
                current_week_file_path = xls_files[0]
                last_week_file_path = xls_files[1]

                compare_neotech_drag(current_week_file_path, last_week_file_path)

                updated_filename = os.path.basename(current_week_file_path).replace('.xls', ' updated.xls')
                updated_path = os.path.join(os.path.dirname(current_week_file_path), updated_filename)
                os.rename(current_week_file_path, updated_path)

                messagebox.showinfo("Success", f"Files compared and updated. New file: {updated_filename}")

            else:
                messagebox.showerror("Error", "At least two .xls files should be in the dropped folder!")
        else:
            messagebox.showerror("Error", "Please drop a folder containing the .xls files!")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    root.title("Drag and Drop Interface")
    root.geometry("400x400")

    app = DropApp(root)

    root.mainloop()
