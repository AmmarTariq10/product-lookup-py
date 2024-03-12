import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime
import os
import requests
import base64
import time
from Crypto.Signature import pkcs1_15
from Crypto.Hash import SHA256
from Crypto.PublicKey import RSA

    

class FilePickerApp:
    def __init__(self, master):
        self.master = master
        self.master.title("File Picker App")
        self.pick_button = tk.Button(self.master, text="Open File Picker", command=self.open_file_picker)
        self.pick_button.pack(pady=20)

    def open_file_picker(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Sheet", "*.xlsx")])
        if file_path:
            print(f"Selected file: {file_path}")
            self.process_and_save_excel(file_path)

    def process_and_save_excel(self, input_file_path, output_file_path=None):
        try:
            df = pd.read_excel(input_file_path)
            df = df.applymap(str)
            # Split UPCs into batches of 20
            upc_batches = [df['UPC'][i:i+20] for i in range(0, len(df['UPC']), 20)]

            # Fetch data for each batch and concatenate the results
            result_data = pd.concat([self.get_walmart_data_batch(batch) for batch in upc_batches], ignore_index=True)

            # Convert integer values in result_data to strings
            # result_data = result_data.applymap(str)

            # Merge the result data with the original DataFrame
            df = pd.concat([df, result_data], axis=1)

            # Save the updated DataFrame to the new Excel file
            if output_file_path is None:
                file_name, file_extension = os.path.splitext(os.path.basename(input_file_path))
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                output_file_path = os.path.join(
                    os.path.dirname(input_file_path),
                    f"{file_name}-processed-at-{timestamp}{file_extension}"
                )
            df.to_excel(output_file_path, index=False)
            print(f"Updated data saved to {output_file_path}")
        except Exception as e:
            print(f"Error processing Excel file: {e}")
            
    def generate_signature(self, key_path, string_to_sign):
            key = self.read_private_key(key_path)
            private_key = RSA.import_key(key)
            h = SHA256.new(string_to_sign.encode('utf-8'))
            signature = pkcs1_15.new(private_key).sign(h)
            signature_string = base64.b64encode(signature).decode('utf-8')
            return signature_string
        
    def read_private_key(self,file_path):
        with open(file_path, 'r') as file:
            private_key_content = file.read()
        return private_key_content
    
    def canonicalize(self,headers_to_sign):
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