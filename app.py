import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from datetime import datetime
import os
import requests
import base64
import time
from Crypto.Signature import pkcs1_15
from Crypto.Hash import SHA256
from Crypto.PublicKey import RSA
import os
import platform
import subprocess

# import pdb; pdb.set_trace()
class FilePickerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Walmart Product Lookup")

        # Get screen width and height
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()

        # Set window width and height to 40% of screen
        window_width = int(screen_width * 0.4)
        window_height = int(screen_height * 0.4)

        # Set window position to center of the screen
        x_position = (screen_width - window_width) // 2
        y_position = (screen_height - window_height) // 2

        # Configure window size and position
        self.master.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

        # Main frame
        self.main_frame = tk.Frame(self.master)
        self.main_frame.pack(expand=True, fill="both")

        # Title and subtitle
        title_label = tk.Label(self.main_frame, text="Walmart Product Lookup", font=("Helvetica", 18, "bold"))
        title_label.pack(pady=10)

        subtitle_label = tk.Label(self.main_frame, text="Select a file to proceed", font=("Helvetica", 14))
        subtitle_label.pack(pady=10)

        # File picker button
        self.pick_button = tk.Button(self.main_frame, text="Open File Picker", command=self.open_file_picker)
        self.pick_button.pack(pady=20)

        # Completion percentage label
        self.completion_label = tk.Label(self.main_frame, text="", font=("Helvetica", 12))

    def open_file_picker(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Sheet", "*.xlsx")])
        if file_path:
            print(f"Selected file: {file_path}")
            # Disable the file picker button during processing
            self.pick_button.config(state=tk.DISABLED)
            # Show the completion label
            self.completion_label.pack(pady=10)
            self.process_and_save_excel(file_path)

    def process_and_save_excel(self, input_file_path, output_file_path=None):
        try:
            df = pd.read_excel(input_file_path)
            df = df.applymap(str)
            # Split UPCs into batches of 20
            upc_batches = [df['UPC'][i:i + 20] for i in range(0, len(df['UPC']), 20)]

            # Completion percentage initialization
            total_batches = len(upc_batches)
            completed_batches = 0

            for i, batch in enumerate(upc_batches):
                result_data = self.get_walmart_data_batch(batch)
                df = pd.concat([df, result_data], axis=1)
                completed_batches += 1
                completion_percentage = (completed_batches / total_batches) * 100
                self.update_completion(completion_percentage)

            # Merge the result data with the original DataFrame
            self.complete_completion()
            self.prompt_completion(df, input_file_path)

        except Exception as e:
            print(f"Error processing Excel file: {e}")
        finally:
            # Enable the file picker button after processing is complete
            self.pick_button.config(state=tk.NORMAL)

    def update_completion(self, completion_percentage):
        self.completion_label.config(text=f"Completion: {completion_percentage:.2f}%")
        self.master.update_idletasks()

    def complete_completion(self):
        self.completion_label.config(text="Completion: 100%")
        self.master.update_idletasks()

    def prompt_completion(self, df, input_file_path):
        response = messagebox.askyesno("Processing Complete", "Do you want to open the processed file?")
        if response:
            self.open_processed_file(df, input_file_path)

    def open_processed_file(self, df, input_file_path):
        # Save the updated DataFrame to the new Excel file
        file_name, file_extension = os.path.splitext(os.path.basename(input_file_path))
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        output_file_path = os.path.join(
            os.path.dirname(input_file_path),
            f"{file_name}-processed-at-{timestamp}{file_extension}"
        )
        df.to_excel(output_file_path, index=False)
        print(f"Updated data saved to {output_file_path}")

        # Open the processed file
        self.open_file(output_file_path)

    def open_file(self, file_path):
        system_platform = platform.system()
        if system_platform == 'Windows':
            os.system(f'start excel "{file_path}"')
        elif system_platform == 'Darwin':  # macOS
            subprocess.run(['open', file_path])
        elif system_platform == 'Linux':
            subprocess.run(['xdg-open', file_path])

    def generate_signature(self, key_path, string_to_sign):
        key = self.read_private_key(key_path)
        private_key = RSA.import_key(key)
        h = SHA256.new(string_to_sign.encode('utf-8'))
        signature = pkcs1_15.new(private_key).sign(h)
        signature_string = base64.b64encode(signature).decode('utf-8')
        return signature_string

    def read_private_key(self, file_path):
        with open(file_path, 'r') as file:
            private_key_content = file.read()
        return private_key_content

    def canonicalize(self, headers_to_sign):
        canonicalized_str_buffer = []
        parameter_names_buffer = []

        sorted_key_set = sorted(headers_to_sign.keys())

        for key in sorted_key_set:
            val = headers_to_sign[key]
            parameter_names_buffer.append(f"{key.strip()};")
            canonicalized_str_buffer.append(f"{val.strip()}\n")

        return ["".join(parameter_names_buffer), "".join(canonicalized_str_buffer)]

    def generate_headers(self, consumer_id, private_key_path, private_key_version):
        intimestamp = str(int(time.time() * 1000))
        headers_map = {
            "WM_CONSUMER.ID": consumer_id,
            "WM_CONSUMER.INTIMESTAMP": intimestamp,
            "WM_SEC.KEY_VERSION": private_key_version
        }
        parameter_names, canonicalized_str = self.canonicalize(headers_map)
        signature = self.generate_signature(private_key_path, canonicalized_str)

        return {
            **headers_map,
            'WM_SEC.AUTH_SIGNATURE': signature
        }

    def get_walmart_data_batch(self, upc_batch):
        # Replace the following with the actual base URL and headers for the Walmart API
        base_url = "https://developer.api.walmart.com/api-proxy/service/affil/product/v2/items"
        headers = self.generate_headers("cc618082-c19d-4c2c-9786-5bfd9e7e0219", "./privateKey.pem", "1")

        # Make the API request here for the entire batch
        api_url = f"{base_url}?upc={','.join(upc_batch)}"
        response = requests.get(api_url, headers=headers)
        print(response.json())
        if response.status_code == 200:
            # Extract relevant information from the response JSON
            result_data = self.extract_walmart_data(response.json())
            return result_data
        else:
            # If there's an error, return a DataFrame with "not found" for each entry
            return pd.DataFrame({"Name": ["not found"] * len(upc_batch)})

    def extract_walmart_data(self, json_response):
        items = json_response.get("items", [])
        result_data = pd.DataFrame({
            "Name": [item.get("name", "not found") for item in items],
            # Add other fields as needed
        })
        return result_data


def main():
    root = tk.Tk()
    app = FilePickerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
