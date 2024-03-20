import os
import sys
import subprocess
import tkinter as tk
import RVToolsAnalysis

class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("500x380")
        #self.root.wm_maxsize(width=600, height=0)
        self.root.title("OPPC Analyzer")

        # Create a label widget
        self.label = tk.Label(self.root, text="This Toolkit analyzes data from RVTools, Nutanix Collector, and LiveOptics")
        self.label.pack(padx=10, pady=10)

        # Create RVTools Button
        self.button1 = tk.Button(self.root, text="RVTools Analysis", command=lambda: self.on_click(self.button1))
        self.button1.pack()

        # Create LiveOptics Button
        self.button2 = tk.Button(self.root, text="LiveOptics Analysis", command=lambda: self.on_click(self.button2))
        self.button2.pack()

        # Create Nutanix Collector Button
        self.button3 = tk.Button(self.root, text="Nutanix Collector Analysis", command=lambda: self.on_click(self.button3))
        self.button3.pack()

        # Create Exit Button
        self.button4 = tk.Button(self.root, text="Exit", command=lambda: self.on_click(self.button4))
        self.button4.pack(padx=0, pady=10)

        # Create Status Label
        self.label1 = tk.Label(self.root, text="Status:")
        self.label1.pack(padx=0, pady=10)

        # Create Status Label
        self.label2 = tk.Label(self.root, text="Please make a selection above to begin", bd=1, width=60, height=10, wraplength=350, relief="sunken")
        self.label2.pack()

        # Start the main event loop
        self.root.mainloop()

    def on_click(self, button):
        # Determine which button was clicked
        if button == self.button1:
            try:
                self.label2.config(text="RVTools Analysis Started\n \nPlease Wait...")
                #self.label2.update_idletasks()  # Force the label update
                #RVToolsAnalysis.main(self.label2)  # Pass label2 as an argument
                def start_rvtoolsanalysis():
                    RVToolsAnalysis.main(self.label2) 
                self.root.after(10, start_rvtoolsanalysis)  # Adjust the '10' for delay as needed
            except Exception as e:
                self.label2.config(text=f"Error: {str(e)}")
        elif button == self.button2:
            self.label2.config(text="LiveOptics Analysis Coming Soon!")
        elif button == self.button3:
            self.label2.config(text="Nutanix Collector Analysis Coming Soon!")
        elif button == self.button4:
            self.root.destroy()
            sys.exit()
        else:
            self.label2.config(text="Unknown button clicked")

app = App()