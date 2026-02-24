import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

def generate_report():
    # 1. Select the CSV File
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    if not file_path:
        return

    try:
        # 2. Load Data & Create Stats
        df = pd.read_csv(file_path)
        
        # Create a Pivot Table (Sales by Category)
        pivot = df.groupby('Category')['Sales'].sum()
        
        # 3. Create a Chart using Matplotlib
        plt.figure(figsize=(6,4))
        pivot.plot(kind='bar', color='skyblue')
        plt.title('Total Sales by Category')
        plt.ylabel('Sales ($)')
        plt.tight_layout()
        plt.savefig('chart.png') # Save temp image
        
        # 4. Export to Excel
        output_path = "Business_Report.xlsx"
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Raw Data', index=False)
            pivot.to_excel(writer, sheet_name='Summary')

        # 5. Add the Chart to Excel using Openpyxl
        wb = load_workbook(output_path)
        ws = wb['Summary']
        img = Image('chart.png')
        ws.add_image(img, 'E2') # Place chart at cell E2
        wb.save(output_path)

        messagebox.showinfo("Success", f"Report saved as {output_path}")
        
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# --- GUI Setup ---
root = tk.Tk()
root.title("Excel Report Generator")
root.geometry("400x200")

label = tk.Label(root, text="CSV to Excel Report Tool", font=("Arial", 14))
label.pack(pady=20)

gen_button = tk.Button(root, text="Upload CSV & Generate", command=generate_report, 
                       bg="#4CAF50", fg="white", font=("Arial", 10, "bold"))
gen_button.pack(pady=10)

root.mainloop()