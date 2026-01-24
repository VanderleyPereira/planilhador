import tkinter as tk
from src.icon.icon import ICONE_BASE64

root = tk.Tk()
img = tk.PhotoImage(data=ICONE_BASE64)
label = tk.Label(root, image=img)
label.pack()
root.mainloop()
