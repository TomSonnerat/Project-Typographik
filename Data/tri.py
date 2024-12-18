import os
import pandas as pd
import tkinter as tk
from PIL import Image, ImageTk
from PIL.Image import Resampling

data = pd.read_excel('switch_data.xlsx')

window = tk.Tk()
window.title("Keyboard Switch Viewer")

current_index = 0

def erase_switch(event):
    global current_index, data
    
    if current_index > 0:
        image_path = data.iloc[current_index - 1]['ImagePath']
        
        try:
            os.remove(image_path)
        except FileNotFoundError:
            print(f"Image file not found: {image_path}")
        except PermissionError:
            print(f"Permission denied to delete: {image_path}")
        except Exception as e:
            print(f"Error deleting image: {e}")
        
        data = data.drop(data.index[current_index - 1])
        data.to_excel('switch_data.xlsx', index=False)
    
    display_next_switch()

def display_next_switch():
    global current_index
    
    if current_index < len(data):
        switch_name = data.iloc[current_index]['Name']
        image_path = data.iloc[current_index]['ImagePath']
        
        image = Image.open(image_path)
        image = image.resize((200, 200), Resampling.BICUBIC)
        photo = ImageTk.PhotoImage(image)
        image_label.configure(image=photo)
        image_label.image = photo
        
        name_label.configure(text=switch_name)
        
        window.bind('1', keep_switch)
        window.bind('2', remove_switch)
        window.bind('3', erase_switch)
        
        current_index += 1
    else:
        window.destroy()

def keep_switch(event):
    display_next_switch()

def remove_switch(event):
    display_next_switch()

image_label = tk.Label(window)
image_label.pack(side=tk.TOP, padx=10, pady=10)

name_label = tk.Label(window, font=('Arial', 16, 'bold'))
name_label.pack(side=tk.TOP, padx=10, pady=10)

instructions_label = tk.Label(window, text="1: Keep Switch, 2: Remove Switch, 3: Erase Photo and Row")
instructions_label.pack(side=tk.BOTTOM, padx=10, pady=10)

display_next_switch()

window.mainloop()