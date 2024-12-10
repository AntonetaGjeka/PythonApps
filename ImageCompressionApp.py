import os #Imports the os module for interacting with the operating system (like file directories)
import tkinter as tk #Imports the tkinter library to create a GUI application
from tkinter import filedialog, messagebox #imports specific dialog-related modules from tkinter
from PIL import Image #Imports the Image class from the PIL (Pillow) library to work with images
import matplotlib.pyplot as plt #Imports matlabs pyplot library to create graphs

#Class definition for the picture compressor GUI application
class PictureCompressor:
    def __init__(self, root):

        #Initializes the GUI application and sets the window title
        self.root = root
        self.root.title("Picture Compressor")

        #Variables to hold the size of the images before and after compression
        self.original_size = 0
        self.compressed_size = 0

        #Create and pack the folder path label
        self.folder_path_label = tk.Label(root, text="Select a folder:") # Label prompting user to select a folder
        self.folder_path_label.pack() # Adds the label to the window

        #Create and pack the entry field where the selected folder path will be displayed
        self.folder_path_entry = tk.Entry(root, width=50) # Entry field for the folder path (50 chars wide)
        self.folder_path_entry.pack() # Adds the entry field to the window

        # Create and pack the "Browse" button to open file dialog for selecting a folder
        self.browse_button = tk.Button(root, text="Browse", command=self.select_folder) # Button to open folder dialog
        self.browse_button.pack() # Adds the button to the window
        
        # Create and pack the "Compress Pictures" button to start the compression process
        self.compress_button = tk.Button(root, text="Compress Pictures", command=self.compress_pictures) # Button to start compression
        self.compress_button.pack() # Adds the button to the window

        # Create and pack the "Show Graph" button to display the graph comparing original and compressed sizes
        self.graph_button = tk.Button(root, text="Show Graph", command=self.show_graph) # Button to show the graph
        self.graph_button.pack() # Adds the button to the window

        # Create and pack the status label to show the compression status
        self.status_label = tk.Label(root, text="Compression status: Not started") # Status label
        self.status_label.pack() # Adds the label to the window

    # Function that opens a folder selection dialog and inserts the folder path into the entry field
    def select_folder(self):
        folder_path = filedialog.askdirectory()  # Opens folder dialog to select a folder
        self.folder_path_entry.delete(0, tk.END)  # Clears any existing path in the entry field
        self.folder_path_entry.insert(0, folder_path)  # Inserts the selected folder path into the entry field


    # Function to compress pictures in the selected folder
    def compress_pictures(self):
        folder_path = self.folder_path_entry.get()  # Gets the folder path from the entry field
        # List comprehension to get all image files in the folder with specified extensions
        pictures = [f for f in os.listdir(folder_path) if f.endswith(('.jpg', '.jpeg', '.png', '.gif', '.bmp'))]
        # Opens a folder dialog to select where to save the compressed pictures
        save_folder_path = filedialog.askdirectory()  
        
        # Calculate the total size of original images
        self.original_size = sum(os.path.getsize(os.path.join(folder_path, f)) for f in pictures)
        
        # Iterate over all selected pictures for compression
        for picture in pictures:
            image_path = os.path.join(folder_path, picture)  # Full path of the image
            image = Image.open(image_path)  # Opens the image using PIL
            # Saves the image to the save folder with 50% quality and optimization for compression
            image.save(os.path.join(save_folder_path, picture), quality=50, optimize=True)
        
        # Calculate the total size of the compressed images
        self.compressed_size = sum(os.path.getsize(os.path.join(save_folder_path, f)) for f in pictures)
        
        # Updates the status label to indicate compression is complete
        self.status_label.config(text="Compression status: Completed")

    # Function to show a bar graph comparing original and compressed folder sizes
    def show_graph(self):
        # Calculate the percentage size reduction
        reduction_percentage = ((self.original_size - self.compressed_size) / self.original_size) * 100
        # Create a bar chart with 'Original' and 'Compressed' sizes
        plt.bar(['Original', 'Compressed'], [self.original_size, self.compressed_size])
        # Set the labels and title for the graph
        plt.xlabel('Folder')
        plt.ylabel('Size (bytes)')
        plt.title('Folder Size Before and After Compression\nSize reduced by {:.2f}%'.format(reduction_percentage))
        # Show the graph
        plt.show()

# Main block to create the Tkinter window and run the application
root = tk.Tk()  # Creates the main window
app = PictureCompressor(root)  # Creates an instance of the PictureCompressor class
root.mainloop()  # Starts the Tkinter event loop to keep the window open
