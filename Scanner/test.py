import pytesseract
from tkinter import Tk, filedialog, Button, Label
from PIL import Image, ImageTk
import cv2
import os

# Set the path to Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def preprocess_image(image_path):
    """
    Preprocess the image for better OCR accuracy.
    Steps include grayscale conversion, thresholding, and noise reduction.
    """
    img = cv2.imread(image_path)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    _, thresh_img = cv2.threshold(blurred, 150, 255, cv2.THRESH_BINARY)
    return thresh_img

def ocr_from_image(image_path):
    """
    Extract text from an image using OCR after preprocessing.
    """
    processed_image = preprocess_image(image_path)
    pil_img = Image.fromarray(processed_image)
    text = pytesseract.image_to_string(pil_img, config='--psm 6')
    return text

def save_to_file(image_path, text):
    """
    Save extracted text to a file with the same name as the image file.
    """
    file_name = os.path.splitext(os.path.basename(image_path))[0] + ".txt"
    output_path = os.path.join(os.getcwd(), file_name)
    
    with open(output_path, "w") as f:
        f.write(text)
    
    return output_path

def upload_image():
    """
    Function to upload an image file using the Tkinter file dialog and process it.
    """
    # Open file dialog to choose an image
    image_path = filedialog.askopenfilename(title="Select an Image File", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg;*.bmp;*.gif")])
    
    if image_path:
        # Perform OCR on the uploaded image
        extracted_text = ocr_from_image(image_path)

        # Save extracted text to a file with the image name
        output_file = save_to_file(image_path, extracted_text)

        # Update label to show success message
        label_output.config(text=f"Text extracted and saved to: {output_file}")

        # Optional: Display the uploaded image (for reference)
        img = Image.open(image_path)
        img.thumbnail((200, 200))
        img_tk = ImageTk.PhotoImage(img)
        label_image.config(image=img_tk)
        label_image.image = img_tk  # keep a reference

# Set up the main Tkinter window
root = Tk()
root.title("Image OCR Text Extractor")

# Label to display output message
label_output = Label(root, text="Select an image to extract text.", font=("Arial", 14))
label_output.pack(pady=10)

# Button to upload image
button_upload = Button(root, text="Upload Image", font=("Arial", 14), command=upload_image)
button_upload.pack(pady=10)

# Label to display uploaded image
label_image = Label(root)
label_image.pack(pady=10)

# Start the Tkinter event loop
root.mainloop()
