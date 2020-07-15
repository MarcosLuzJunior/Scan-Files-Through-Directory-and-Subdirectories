#Importing packages
import os
import os.path
import pandas as pd

# Set input directory
input_dir = r'C:\Users\John_Smith\Downloads'

#Create a list to append results to
data = []

# Create dataframe to load information into
df = pd.DataFrame()

# Create a mini function to hyperlink the directories
def make_hyperlink(value):
    url = "{}"
    return '=HYPERLINK("%s", "%s")' % (url.format(value), value)

# Run nested FOR loop to look for file extensions and name content in directories and sub directories
for dirpath, dirnames, filenames in os.walk(input_dir):
    for filename in [f for f in filenames if f.endswith('.pdf') or f.endswith('.docx')]: # Assign the .pdfs, .docxs files to list (filename)
        directory_ = os.path.join(dirpath, filename) # Join the name of the file with its path address
        file_name, extension = os.path.splitext(filename) # Use split to get the name and the extension of the file
        if 'John' in filename: # Use IF statement to find keyword in file name
            data.append((file_name, extension, directory_)) # If condition matches append file_name, extension, and address of the file to list (data)

df = pd.DataFrame(data, columns=('File Name', 'Extension','Path Address')) # Use pandas to create dataframe with the information in the list data
df['Path Address'] = df['Path Address'].apply(lambda x: make_hyperlink(x)) # Create hyperlink of the path addresses and replace the original values
df.to_excel(r'Files in Drive.xlsx', index = False) # Save output to an excel file

print(df.head(100))
