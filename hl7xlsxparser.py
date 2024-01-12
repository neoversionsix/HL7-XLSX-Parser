import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from hl7apy.parser import parse_message

def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            global data  # Declare 'data' as a global variable
            data = pd.read_excel(file_path)
            process_hl7(data, file_path)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

def process_hl7(data, file_path):
    if "HL7" not in data.columns:
        messagebox.showerror("Error", "No 'HL7' column found in the file.")
        return

    meanings = []
    for hl7_message in data["HL7"]:
        hl7_message = reformat_hl7_message(hl7_message)
        try:
            parsed_message = parse_message(hl7_message)
            meaning = extract_human_readable(parsed_message)
            meanings.append(meaning)
        except Exception as e:
            meanings.append(f"Error parsing HL7 message: {e}")

    data['MEANING'] = meanings

# Other functions (reformat_hl7_message, extract_human_readable, get_field_description) remain the same
def reformat_hl7_message(hl7_message):
    # Insert line breaks for proper parsing
    for segment in ['MSH', 'PID', 'OBR', 'OBX', 'ORC', 'PV1', 'PV2']:
        hl7_message = hl7_message.replace(segment, '\r' + segment)
    return hl7_message.lstrip('\r')
    
def extract_human_readable(parsed_message):
    result = []
    for segment in parsed_message.children:
        segment_str = segment.to_er7()
        fields = segment_str.split('|')

        for i, field_value in enumerate(fields, start=1):
            if i < len(fields):  # Ensure field index within range
                field_desc = get_field_description(segment.name, str(i))
                field_value = field_value if field_value else "N/A"
                result.append(f"{segment.name}-{i} {field_desc}: {field_value}")

    return '\n'.join(result)




def get_field_description(segment_name, field_index):
    # Manually defined field descriptions
    descriptions = {
        "MSH-1": "Field Separator",
        "MSH-2": "Encoding Characters",
        "MSH-3": "Sending Application",
        "MSH-4": "Sending Facility",
        "MSH-5": "Receiving Application",
        "MSH-6": "Receiving Facility",
        "MSH-7": "Date/Time of Message",
        "MSH-8": "Security",
        "MSH-9": "Message Type",
        "MSH-10": "Message Control ID",
        "MSH-11": "Processing ID",
        "MSH-12": "Version ID",
        "MSH-13": "Sequence Number",
        "MSH-14": "Continuation Pointer",
        "MSH-15": "Accept Acknowledgement Type",
        "MSH-16": "Application Acknowledgement Type",
        "MSH-17": "Country Code",
        "MSH-18": "Character Set",
        "MSH-19": "Principal Language of Message",
        # Add more descriptions as needed
    }
    return descriptions.get(f"{segment_name}-{field_index}", "")



def save_as():
    global data
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            data.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "File saved successfully!")
            root.destroy()  # This will close the program
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

root = tk.Tk()
root.title("HL7 Parser")
root.geometry("300x150")

load_button = tk.Button(root, text="Load Excel File", command=load_file)
load_button.pack(expand=True)

save_button = tk.Button(root, text="Save Excel File", command=save_as)
save_button.pack(expand=True)

root.mainloop()
