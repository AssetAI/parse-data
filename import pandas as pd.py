import pandas as pd
import tkinter as tk
from tkinter import filedialog
from fastkml import kml
from shapely.geometry import Point

# Create a root window and hide it
root = tk.Tk()
root.withdraw()

# Prompt the user for the file location
print("Please choose the input file.")
file_path = filedialog.askopenfilename()

print("Reading file:", file_path)

# Load the spreadsheet data
df = pd.read_excel(file_path)

# Parse the 'Gsensor' column to get the 'g in z direction' values
df['g_z'] = df['Gsensor'].apply(lambda x: [float(i.split(',')[-1]) for i in x.split('; ')])

# Prompt the user for the output location
print("Please enter the output directory.")
output_dir = input()

# Prompt the user for the output file name
print("Please enter the output file name (without extension).")
output_file = input()

output_path = output_dir + "/" + output_file + ".kml"

print("Output path:", output_path)

# Create a new KML object
k = kml.KML()

# Create a new document
doc = kml.Document()

# Add each row as a placemark
for index, row in df.iterrows():
    # Create a new placemark
    pm = kml.Placemark()

    # Set the placemark's coordinates to the Latitude and Longitude of the row
    pm.geometry = Point(row['Longitude'], row['Latitude'])

    # Add the placemark to the document
    doc.append(pm)

# Add the document to the KML object
k.append(doc)

print("Writing KML file...")

# Write the KML object to a file
with open(output_path, 'w') as f:
    f.write(k.to_string(prettyprint=True))

print("KML file has been saved to:", output_path)
