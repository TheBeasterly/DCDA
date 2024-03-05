import os
import sys
import subprocess
import tkinter as tk
import RVToolsAnalysis

class App:
    def __init__(self):
        self.root = tk.Tk()
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

        # Start the main event loop
        self.root.mainloop()

    def on_click(self, button):
        # Determine which button was clicked
        if button == self.button1:
            # Run the RVToolsAnalysis Python script
            print("RVTools Analysis Started")
            RVToolsAnalysis.main()

            print("RVTools Analysis Completed. Analyzed file with be new file appended with -EDITED and saved in same directory as Source File")
        elif button == self.button2:
            print("LiveOptics Analysis Coming Soon!")
        elif button == self.button3:
            print("Nutanix Collector Analysis Coming Soon!")
        elif button == self.button4:
            self.root.destroy()
            sys.exit()
        else:
            print("Unknown button clicked")

app = App()