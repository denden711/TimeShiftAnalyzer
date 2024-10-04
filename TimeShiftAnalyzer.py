import pandas as pd
import matplotlib.pyplot as plt
from tkinter import filedialog, Button, Label, Entry, Tk, Listbox, MULTIPLE, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import os
import logging
from openpyxl import Workbook

# Configure logging for error messages with UTF-8 encoding to prevent character corruption
log_handler = logging.FileHandler('error_log.txt', encoding='utf-8')
log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger = logging.getLogger()
logger.setLevel(logging.ERROR)
logger.addHandler(log_handler)

# Clean and normalize file paths
def clean_file_path(file_path):
    return os.path.normpath(file_path.replace("{", "").replace("}", "").strip())

# Load CSV file with error handling
def load_csv(file_path, encoding='shift-jis'):
    try:
        return pd.read_csv(clean_file_path(file_path), encoding=encoding)
    except Exception as e:
        error_message = f"Failed to load CSV file: {file_path}. Error: {str(e)}"
        logger.error(error_message)
        raise ValueError(error_message)

# Extract relevant columns (Time and Voltage) from the DataFrame
def extract_columns(df, time_col_index, voltage_col_index):
    try:
        return pd.DataFrame({
            'Time': df.iloc[:, time_col_index],
            'Voltage': df.iloc[:, voltage_col_index]
        })
    except IndexError as e:
        error_message = f"Invalid column indices. Error: {str(e)}"
        logger.error(error_message)
        raise ValueError(error_message)

# Filter data based on a time range
def filter_time_range(df, min_time, max_time):
    try:
        return df[(df['Time'] >= min_time) & (df['Time'] <= max_time)]
    except Exception as e:
        error_message = f"Failed to filter data by time range. Error: {str(e)}"
        logger.error(error_message)
        raise ValueError(error_message)

# Find the time when the voltage is closest to zero
def find_zero_point(df):
    try:
        return df.loc[df['Voltage'].abs().idxmin()]['Time']
    except Exception as e:
        error_message = f"Failed to find zero point. Error: {str(e)}"
        logger.error(error_message)
        raise ValueError(error_message)

# Shift the time in the DataFrame by a given time shift
def shift_time(df, time_shift):
    try:
        shifted_df = df.copy()
        shifted_df['Time'] += time_shift
        return shifted_df
    except Exception as e:
        error_message = f"Failed to shift time. Error: {str(e)}"
        logger.error(error_message)
        raise ValueError(error_message)

# Plot the baseline and shifted data
def plot_data(df1, shifted_dfs):
    plt.figure(figsize=(10, 6))
    plt.plot(df1['Time'], df1['Voltage'], label='Baseline CSV - Voltage vs Time', marker='o')
    for i, df_shifted in enumerate(shifted_dfs):
        plt.plot(df_shifted['Time'], df_shifted['Voltage'], label=f'Shifted CSV {i+1} - Voltage vs Time', marker='x')
    plt.title('Aligned Voltage vs Time')
    plt.xlabel('Time (s)')
    plt.ylabel('Voltage (V)')
    plt.legend()
    plt.grid(True)
    plt.show()

# Save shift data to an Excel file
def save_shift_data_to_excel(shift_data):
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_file:
        df = pd.DataFrame(shift_data, columns=['File Name', 'Zero Point Time (seconds)', 'Time Shift (seconds)'])
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Shift Data', index=False)
                worksheet = writer.sheets['Shift Data']
                for column_cells in worksheet.iter_cols(min_col=2, max_col=3, min_row=2, max_row=worksheet.max_row):
                    for cell in column_cells:
                        cell.number_format = '0.00000000000000000000E+00'
            print(f"Shift data saved to {output_file}")
        except Exception as e:
            error_message = f"Failed to save data to Excel. Error: {str(e)}"
            logger.error(error_message)
            raise ValueError(error_message)

# Execute the shifting and plotting operation
def execute_shift_and_plot(baseline_file, shift_files, min_time, max_time):
    try:
        df_baseline = load_csv(baseline_file)
        df_baseline_extracted = extract_columns(df_baseline, time_col_index=17, voltage_col_index=4)
        filtered_df_baseline = filter_time_range(df_baseline_extracted, min_time, max_time)
        baseline_zero_time = find_zero_point(filtered_df_baseline)

        shifted_dfs = []
        shift_data = [[os.path.basename(baseline_file), baseline_zero_time, 0]]  # Baseline shift is 0

        for shift_file in shift_files:
            df_shift = load_csv(shift_file)
            df_shift_extracted = extract_columns(df_shift, time_col_index=17, voltage_col_index=4)
            filtered_df_shift = filter_time_range(df_shift_extracted, min_time, max_time)
            shift_zero_time = find_zero_point(filtered_df_shift)
            time_shift = baseline_zero_time - shift_zero_time
            shifted_dfs.append(shift_time(filtered_df_shift, time_shift))
            shift_data.append([os.path.basename(shift_file), shift_zero_time, time_shift])
            print(f"Time shift applied to {os.path.basename(shift_file)}: {time_shift} seconds")

        save_shift_data_to_excel(shift_data)
        plot_data(filtered_df_baseline, shifted_dfs)

    except ValueError as e:
        messagebox.showerror("Error", str(e))

# Main GUI setup
def main():
    root = TkinterDnD.Tk()
    root.title("TimeShiftAnalyzer")

    baseline_label = Label(root, text="No baseline file selected")
    baseline_label.pack(pady=5)

    shift_files_listbox = Listbox(root, selectmode=MULTIPLE, height=6)
    shift_files_listbox.pack(pady=5, fill='both')

    time_range_label = Label(root, text="Enter Time Range (min, max)")
    time_range_label.pack(pady=5)

    time_range_min_entry = Entry(root)
    time_range_min_entry.pack(pady=5)
    time_range_min_entry.insert(0, "5e-05")

    time_range_max_entry = Entry(root)
    time_range_max_entry.pack(pady=5)
    time_range_max_entry.insert(0, "8e-05")

    baseline_file = None
    shift_files = []

    def select_baseline_file():
        nonlocal baseline_file
        baseline_file = filedialog.askopenfilename(title="Select Baseline CSV File", filetypes=[("CSV files", "*.csv")])
        if baseline_file:
            baseline_label.config(text=f"Selected: {os.path.basename(baseline_file)}")

    def drop_baseline_file(event):
        nonlocal baseline_file
        files = event.data.split()
        if len(files) > 1:
            messagebox.showerror("Error", "You can only select one baseline CSV file.")
        else:
            baseline_file = files[0]
            baseline_label.config(text=f"Selected: {os.path.basename(baseline_file)}")

    def drop_shift_files(event):
        files = event.data.split()
        for file in files:
            if file not in shift_files:
                shift_files_listbox.insert('end', file)
                shift_files.append(file)

    def remove_selected_files():
        selected_indices = shift_files_listbox.curselection()
        for index in selected_indices[::-1]:
            shift_files_listbox.delete(index)
            del shift_files[index]

    def execute():
        if not baseline_file:
            messagebox.showerror("Error", "Please select a baseline CSV file.")
            return

        if not shift_files:
            messagebox.showerror("Error", "Please select at least one shift CSV file.")
            return

        try:
            min_time = float(time_range_min_entry.get())
            max_time = float(time_range_max_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values for the time range.")
            return

        execute_shift_and_plot(baseline_file, shift_files, min_time, max_time)

    baseline_button = Button(root, text="Select Baseline CSV", command=select_baseline_file)
    baseline_button.pack(pady=5)

    remove_button = Button(root, text="Remove Selected Shift Files", command=remove_selected_files)
    remove_button.pack(pady=5)

    execute_button = Button(root, text="Execute Shift and Plot", command=execute)
    execute_button.pack(pady=20)

    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', drop_shift_files)
    baseline_label.drop_target_register(DND_FILES)
    baseline_label.dnd_bind('<<Drop>>', drop_baseline_file)

    root.mainloop()

if __name__ == "__main__":
    main()
