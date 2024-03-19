# Word Automation Program

This program automates the process of replacing images in Word documents with a specified image. It utilizes the Microsoft Office Interop library to manipulate Word documents.

## Functionality

The program performs the following tasks:

- Opens each Word document in a specified folder.
- Searches for inline images within the document.
- Replaces each inline image with a specified image.
- Saves the modified document.

## Requirements

- Microsoft Office must be installed on the system.
- The Word documents must be in the .doc or .docx format.
- Rerference to Microsoft.office.interop.word 

## Usage

1. **Clone the Repository\Download**: Clone\Download the repository containing the program files.

2. **Build the Program**: Build the program using Visual Studio or any compatible C# compiler.

3. **Set Image and Folder Paths**: Update the `imagePath` and `folderPath` variables in the `Program.cs` file with the paths to your image and folder containing Word documents.

4. **Run the Program**: Execute the program. It will process each Word document in the specified folder, replacing inline images with the specified image.

## Important Note

- The program may not pick up some images if the "Warp Text" format of the image is not set to "In line with Text" in the original Word document. If you need to replace images with a different format. I recommend using Spire.doc.

