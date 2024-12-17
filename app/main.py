import tkinter as tk

def start_gui():
    # Create and run window
    root = tk.Tk()
    root.geometry('1300x800')
    root.title('PPTX Generator')
    root.mainloop()

if __name__ == "__main__":
    print("Welcome to PowerPoint Generator")
    start_gui()
