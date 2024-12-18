import os
import pandas as pd
import tkinter as tk
from PIL import Image, ImageTk
from PIL.Image import Resampling
import tkinter.messagebox
import ast

class SwitchColorListPicker:
    def __init__(self, excel_path='switch_data.xlsx'):
        # Read the Excel file
        self.data = pd.read_excel(excel_path)
        
        # Ensure TransparentLevel and ColorListLength columns are string type
        if 'TransparentLevel' not in self.data.columns:
            self.data['TransparentLevel'] = ''
        if 'ColorListLength' not in self.data.columns:
            self.data['ColorListLength'] = ''
        
        # Convert columns to string type
        self.data['TransparentLevel'] = self.data['TransparentLevel'].astype(str)
        self.data['ColorListLength'] = self.data['ColorListLength'].astype(str)

        self.current_index = 0
        self.transparent_level = None
        self.color_list_length = None

        self.window = tk.Tk()
        self.window.title("Switch Transparency and Color List Length Picker")
        self.window.geometry("1000x900")

        self.main_frame = tk.Frame(self.window)
        self.main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)

        self.name_label = tk.Label(self.main_frame, font=('Arial', 16, 'bold'))
        self.name_label.pack(pady=(0, 10))

        # Image Display
        self.image_frame = tk.Frame(self.main_frame)
        self.image_frame.pack(pady=(0, 20))

        self.canvas = tk.Canvas(self.image_frame, width=600, height=600)
        self.canvas.pack()

        # Transparency Level
        self.transparency_frame = tk.Frame(self.main_frame)
        self.transparency_frame.pack(pady=(10, 10))

        self.transparency_label = tk.Label(self.transparency_frame, text="Select Transparency Level:", font=('Arial', 14, 'bold'))
        self.transparency_label.pack(side=tk.LEFT, padx=(0, 10))

        self.transparency_display = tk.Label(self.transparency_frame, text="Not Selected", 
                                             width=15, height=2, 
                                             relief=tk.RAISED, 
                                             font=('Arial', 12))
        self.transparency_display.pack(side=tk.LEFT)

        # Color List Length
        self.color_list_frame = tk.Frame(self.main_frame)
        self.color_list_frame.pack(pady=(10, 10))

        self.color_list_label = tk.Label(self.color_list_frame, text="Select Color List Length:", font=('Arial', 14, 'bold'))
        self.color_list_label.pack(side=tk.LEFT, padx=(0, 10))

        self.color_list_display = tk.Label(self.color_list_frame, text="Not Selected",
                                           width=15, height=2,
                                           relief=tk.RAISED,
                                           font=('Arial', 12))
        self.color_list_display.pack(side=tk.LEFT)

        # Instructions
        self.instruction_label = tk.Label(self.main_frame, 
                                          text="Transparency (Number keys):\n1 - No Transparency\n2 - Body Transparency\n3 - Full Transparency\n\n" +
                                               "Color List Length (Number keys):\n4 - No Change\n5 - Remove 2nd Color\n6 - Remove 2nd and 3rd Colors\n\n" +
                                               "Navigation:\nEnter - Next Switch\nBackspace - Previous Switch", 
                                          font=('Arial', 10), justify=tk.LEFT)
        self.instruction_label.pack(pady=(10, 20))

        # Navigation Frame
        self.nav_frame = tk.Frame(self.main_frame)
        self.nav_frame.pack(pady=(20, 0))

        self.prev_button = tk.Button(self.nav_frame, text="Previous", command=self.previous_switch)
        self.prev_button.pack(side=tk.LEFT, padx=10)

        self.next_button = tk.Button(self.nav_frame, text="Next", command=self.next_switch)
        self.next_button.pack(side=tk.LEFT, padx=10)

        self.current_image = None
        self.current_pil_image = None

        # Binding number keys for transparency and color list selection
        self.window.bind('1', self.select_transparency_no)
        self.window.bind('2', self.select_transparency_body)
        self.window.bind('3', self.select_transparency_full)

        self.window.bind('4', self.select_color_list_no_change)
        self.window.bind('5', self.select_color_list_remove_second)
        self.window.bind('6', self.select_color_list_remove_second_third)

        # Add Enter key binding for next switch
        self.window.bind('<Return>', self.next_switch)
        
        # Add Backspace key binding for previous switch
        self.window.bind('<BackSpace>', self.previous_switch)

        self.display_current_switch()

    def select_transparency_no(self, event=None):
        self.transparent_level = "No"
        self.transparency_display.configure(text="Transparency: No", bg='lightgreen')

    def select_transparency_body(self, event=None):
        self.transparent_level = "Body"
        self.transparency_display.configure(text="Transparency: Body", bg='yellow')

    def select_transparency_full(self, event=None):
        self.transparent_level = "Full"
        self.transparency_display.configure(text="Transparency: Full", bg='lightblue')

    def select_color_list_no_change(self, event=None):
        self.color_list_length = "NoChange"
        self.color_list_display.configure(text="Color List: No Change", bg='lightgray')

    def select_color_list_remove_second(self, event=None):
        self.color_list_length = "RemoveSecond"
        self.color_list_display.configure(text="Color List: Remove 2nd Color", bg='lightyellow')

        try:
            colors = ast.literal_eval(self.data.iloc[self.current_index]['Colors'])
            if len(colors) > 1:
                colors.pop(1)
                self.data.loc[self.current_index, 'Colors'] = str(colors)
        except Exception as e:
            print(f"Error modifying color list: {e}")

    def select_color_list_remove_second_third(self, event=None):
        self.color_list_length = "RemoveSecondThird"
        self.color_list_display.configure(text="Color List: Remove 2nd and 3rd Colors", bg='lightcoral')

        try:
            colors = ast.literal_eval(self.data.iloc[self.current_index]['Colors'])
            while len(colors) > 1:
                colors.pop(1)
            self.data.loc[self.current_index, 'Colors'] = str(colors)
        except Exception as e:
            print(f"Error modifying color list: {e}")

    def display_current_switch(self):
        # Reset selections
        self.transparent_level = None
        self.color_list_length = None

        # Reset displays
        self.transparency_display.configure(text="Transparency: Not Selected", bg='SystemButtonFace')
        self.color_list_display.configure(text="Color List: Not Selected", bg='SystemButtonFace')

        switch = self.data.iloc[self.current_index]

        self.name_label.configure(text=switch['Name'])

        try:
            self.current_pil_image = Image.open(switch['ImagePath'])
            self.current_pil_image.thumbnail((600, 600), Resampling.LANCZOS)
            self.current_image = ImageTk.PhotoImage(self.current_pil_image)
            self.canvas.delete('all')
            self.canvas.create_image(300, 300, image=self.current_image, anchor=tk.CENTER)
        except Exception as e:
            print(f"Error loading image: {e}")
            self.canvas.delete('all')
            self.canvas.create_text(300, 300, text="Image not available", font=('Arial', 14), fill='red', anchor=tk.CENTER)

    def next_switch(self, event=None):
        print(f"Current transparency level: {self.transparent_level}")
        print(f"Current color list length: {self.color_list_length}")

        if self.transparent_level is None:
            tk.messagebox.showwarning("Warning", "Please select Transparency Level!")
            return

        if self.color_list_length is None:
            tk.messagebox.showwarning("Warning", "Please select Color List Length!")
            return

        self.save_selections()

        if self.current_index < len(self.data) - 1:
            self.current_index += 1
            self.display_current_switch()
        else:
            tk.messagebox.showinfo("Information", "You have reached the last switch.")
            self.window.quit()

    def previous_switch(self, event=None):
        if self.transparent_level is None:
            tk.messagebox.showwarning("Warning", "Please select Transparency Level!")
            return

        if self.color_list_length is None:
            tk.messagebox.showwarning("Warning", "Please select Color List Length!")
            return

        self.save_selections()

        if self.current_index > 0:
            self.current_index -= 1
            self.display_current_switch()
        else:
            tk.messagebox.showinfo("Information", "You have reached the first switch.")

    def save_selections(self):
        if self.transparent_level and self.color_list_length:
            self.data.loc[self.current_index, 'TransparentLevel'] = self.transparent_level
            self.data.loc[self.current_index, 'ColorListLength'] = self.color_list_length
            self.data.to_excel('switch_data.xlsx', index=False)

            print(f"Saved for {self.data.iloc[self.current_index]['Name']}: "
                  f"Transparency: {self.transparent_level}, "
                  f"Color List:")

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = SwitchColorListPicker()
    app.run()