import os
import pandas as pd
import tkinter as tk
from PIL import Image, ImageTk
from PIL.Image import Resampling
import numpy as np

class SwitchColorPicker:
    def __init__(self, excel_path='switch_data.xlsx'):
        self.data = pd.read_excel(excel_path)
        
        self.current_index = 0
        
        self.selected_colors = []
        
        self.window = tk.Tk()
        self.window.title("Switch Image Color Picker")
        self.window.geometry("1000x900")
        
        self.main_frame = tk.Frame(self.window)
        self.main_frame.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
        
        self.name_label = tk.Label(self.main_frame, font=('Arial', 16, 'bold'))
        self.name_label.pack(pady=(0,10))
        
        self.image_color_frame = tk.Frame(self.main_frame)
        self.image_color_frame.pack(pady=(0,20))
        
        self.canvas = tk.Canvas(self.image_color_frame, width=600, height=600, cursor="cross")
        self.canvas.pack(side=tk.LEFT, padx=(0,20))
        
        self.zoom_canvas = tk.Canvas(self.image_color_frame, width=200, height=200)
        self.zoom_canvas.pack(side=tk.LEFT)
        
        self.pixel_color_label = tk.Label(self.main_frame, text="Pixel Color: ", font=('Arial', 12), width=30)
        self.pixel_color_label.pack(pady=(10,0))
        
        self.canvas.bind('<Button-1>', self.pick_color)
        self.canvas.bind('<Motion>', self.update_zoom_and_color)
        
        self.color_frame = tk.Frame(self.main_frame)
        self.color_frame.pack(pady=(20,20))
        
        self.color_displays = []
        for i in range(3):
            color_display = tk.Label(self.color_frame, text=f"Color {i+1}", width=20, height=2, relief=tk.RAISED)
            color_display.grid(row=0, column=i, padx=10)
            self.color_displays.append(color_display)
        
        self.nav_frame = tk.Frame(self.main_frame)
        self.nav_frame.pack(pady=(20,0))
        
        self.prev_button = tk.Button(self.nav_frame, text="Previous", command=self.previous_switch)
        self.prev_button.pack(side=tk.LEFT, padx=10)
        
        self.next_button = tk.Button(self.nav_frame, text="Next", command=self.next_switch)
        self.next_button.pack(side=tk.LEFT, padx=10)
        
        self.save_button = tk.Button(self.nav_frame, text="Save Colors", command=self.save_colors)
        self.save_button.pack(side=tk.LEFT, padx=10)
        
        self.current_image = None
        self.current_pil_image = None
        self.current_pixel_color = None
        
        self.window.bind('1', self.select_color_1)
        self.window.bind('2', self.select_color_2)
        self.window.bind('3', self.select_color_3)
        
        self.last_x = None
        self.last_y = None
        
        self.display_current_switch()
        
    def get_average_color(self, x, y):
        colors = []
        for dx in [-1, 0, 1]:
            for dy in [-1, 0, 1]:
                sample_x = x + dx
                sample_y = y + dy
                
                if (0 <= sample_x < self.current_pil_image.width and 
                    0 <= sample_y < self.current_pil_image.height):
                    color = self.current_pil_image.getpixel((sample_x, sample_y))
                    colors.append(color)
        
        avg_color = tuple(map(lambda x: int(sum(x)/len(x)), zip(*colors)))
        return avg_color
    
    def select_color_1(self, event=None):
        if self.current_pixel_color:
            self.select_color(0, self.current_pixel_color)
    
    def select_color_2(self, event=None):
        if self.current_pixel_color:
            self.select_color(1, self.current_pixel_color)
    
    def select_color_3(self, event=None):
        if self.current_pixel_color:
            self.select_color(2, self.current_pixel_color)
    
    def select_color(self, index, color):
        if index < 3:
            self.color_displays[index].configure(bg=color, text=color)
            
            if index < len(self.selected_colors):
                self.selected_colors[index] = color
            else:
                self.selected_colors.append(color)
            
            if len(self.selected_colors) == 3:
                self.next_switch()
    
    def display_current_switch(self):
        self.selected_colors = []
        
        for label in self.color_displays:
            label.configure(bg='SystemButtonFace', text="Click to pick color")
        
        self.pixel_color_label.configure(text="Pixel Color: ")
        
        switch = self.data.iloc[self.current_index]
        
        self.name_label.configure(text=switch['Name'])
        
        self.current_pil_image = Image.open(switch['ImagePath'])
        
        self.current_pil_image.thumbnail((600, 600), Resampling.LANCZOS)
        
        self.current_image = ImageTk.PhotoImage(self.current_pil_image)
        
        self.canvas.delete('all')
        self.canvas.create_image(300, 300, image=self.current_image, anchor=tk.CENTER)
        
        self.zoom_canvas.delete('all')
        
        if pd.notna(switch['Colors']):
            existing_colors = eval(switch['Colors'])
            for i, color in enumerate(existing_colors):
                if i < 3:
                    self.color_displays[i].configure(bg=color, text=color)
                    self.selected_colors.append(color)
        
    def update_zoom_and_color(self, event):
        if self.current_pil_image is None:
            return
        
        x = event.x - (300 - self.current_pil_image.width // 2)
        y = event.y - (300 - self.current_pil_image.height // 2)
        
        self.last_x = x
        self.last_y = y
        
        if 0 <= x < self.current_pil_image.width and 0 <= y < self.current_pil_image.height:
            avg_color = self.get_average_color(x, y)
            hex_color = '#{:02x}{:02x}{:02x}'.format(*avg_color)
            
            self.current_pixel_color = hex_color
            
            self.pixel_color_label.configure(text=f"Pixel Color: {hex_color} RGB{avg_color}")
            
            self.zoom_canvas.delete('all')
            
            zoom_size = 10
            
            left = max(0, x - zoom_size // 2)
            top = max(0, y - zoom_size // 2)
            right = min(self.current_pil_image.width, left + zoom_size)
            bottom = min(self.current_pil_image.height, top + zoom_size)
            
            zoom_img = self.current_pil_image.crop((left, top, right, bottom))
            zoom_img = zoom_img.resize((200, 200), Resampling.NEAREST)
            
            zoom_photo = ImageTk.PhotoImage(zoom_img)
            
            self.zoom_canvas.create_image(100, 100, image=zoom_photo, anchor=tk.CENTER)
            self.zoom_canvas.zoom_photo = zoom_photo
            
            self.zoom_canvas.create_line(0, 100, 200, 100, fill='red')
            self.zoom_canvas.create_line(100, 0, 100, 200, fill='red')
    
    def pick_color(self, event):
        if self.current_pixel_color:
            current_index = len(self.selected_colors)
            self.select_color(current_index, self.current_pixel_color)
    
    def next_switch(self):
        self.save_colors()
        
        if self.current_index < len(self.data) - 1:
            self.current_index += 1
            self.display_current_switch()
        else:
            self.window.quit()
    
    def previous_switch(self):
        self.save_colors()
        
        if self.current_index > 0:
            self.current_index -= 1
            self.display_current_switch()
    
    def save_colors(self):
        if self.selected_colors:
            self.data.loc[self.current_index, 'Colors'] = str(self.selected_colors)
            
            self.data.to_excel('switch_data.xlsx', index=False)
            
            print(f"Saved colors for {self.data.iloc[self.current_index]['Name']}: {self.selected_colors}")
    
    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = SwitchColorPicker()
    app.run()