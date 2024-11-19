import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd 
import json 
import os


def upload_file():
    #open file dialog to selct the excel/csv file
    file_path = filedialog.askopenfilename(
        title="Select the excel/csv file",
        filetypes=[("Excel files","*.xlsx *xls"),("CSV Files","*.csv")]
    )

    if file_path:
        try:
            #read the file into a pandas data frame
            if file_path.endswith(".csv"):
                data = pd.read_csv(file_path)
            else:
                data = pd.read_excel(file_path)


                #call function to convert dataframe to json
                convert_to_json(data,file_path)
        except Exception as e:
            messagebox.showerror("Error",f"Failed to process the file :{e}")



def convert_to_json(data , orginal_file_path):
    json_file_path = orginal_file_path.rsplit('.',1)[0] + ".json"

    try:
        #convert the dataframe to json and save it 
        data.to_json(json_file_path,orient="records",indent=4)

        #show success message 
        messagebox.showinfo("Success",f"File Converted and saved as {json_file_path}")

        #Ask user if they want to open the json file 
        if messagebox.askyesno("Open file","Do youu want to open the json file?"):
            open_json_file(json_file_path)

    except Exception as e:
        messagebox.showerror("Error",f"Failed to convert to JSON:{e}")


#open json file
def open_json_file(file_path):
    try:
        #open the json file with default application 
        os.startfile(file_path)
    except Exception as e:
        messagebox.showerror("Error: ",f"failed to open the file : {e}")





#creating the main window 

def create_window():
    #create window 
    window = tk.Tk()
    window.title("Excel to Json Convertor")
    window.geometry("400x200")

    #add label 
    label = tk.Label(window , text ="Excel/csv to json convertor",font =("Arial",16))
    label.pack(pady=10)

    #add a button to the window to upload files
    upload_button = tk.Button(window,text ="Upload Excel / CSV file",command = upload_file)
    upload_button.pack(pady=10)

    #run the window on loop until the user closes it .
    window.mainloop()


if __name__ == "__main__":
    create_window()
