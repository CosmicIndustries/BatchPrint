#for windows os
import os
import time
import win32print
import win32ui
from PIL import Image, ImageWin

def print_file_with_default_app(file_path):
    """Print a .docx file using the system's default app."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.docx':
        try:
            print(f"Printing document: {file_path}")
            os.startfile(file_path, "print")
            time.sleep(5)
            print("Print command sent successfully.")
        except Exception as e:
            print(f"An error occurred while printing document: {e}")
    elif ext == '.pdf':
        print(f"No handler for PDF printing: {file_path}")
        print("Hint: Associate a PDF reader (like Adobe Acrobat) with .pdf files or use win32print helpers.")

def print_image_win32(image_path):
    """Send an image to the default printer using Win32 API."""
    try:
        printer_name = win32print.GetDefaultPrinter()
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)

        img = Image.open(image_path)
        dib = ImageWin.Dib(img)

        hDC.StartDoc(os.path.basename(image_path))
        hDC.StartPage()
        dib.draw(hDC.GetHandleOutput(), (0, 0, img.width, img.height))
        hDC.EndPage()
        hDC.EndDoc()
        hDC.DeleteDC()

        print(f"Successfully printed: {image_path}")
    except Exception as e:
        print(f"Failed to print image: {image_path}, Error: {e}")

def batch_print_folder(folder_path):
    """Print all .docx, .pdf (warn), and image files in the given folder."""
    if not os.path.isdir(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist.")
        return

    files = os.listdir(folder_path)
    doc_files = [f for f in files if f.lower().endswith(('.docx', '.pdf'))]
    image_files = [f for f in files if f.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))]

    for doc in doc_files:
        file_path = os.path.join(folder_path, doc)
        print_file_with_default_app(file_path)

    for image in image_files:
        image_path = os.path.join(folder_path, image)
        print_image_win32(image_path)

    print("\nBatch print attempt completed.")

if __name__ == "__main__":
    folder = input("Enter the folder path containing documents and images to print: ").strip()
    batch_print_folder(folder)

