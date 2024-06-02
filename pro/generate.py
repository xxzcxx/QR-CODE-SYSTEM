import tkinter as tk
from tkinter import messagebox, filedialog
import sqlite3
import qrcode
from PIL import Image, ImageTk
import datetime
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from PIL import Image as PILImage
import shutil
import os
class UserRegistration:
    
    def __init__(self, root):
        self.root = root
        self.root.title("User Registration")
        
        root.geometry("600x700")  # Adjust the window size as needed
        self.name_var = tk.StringVar()
        self.stage_var = tk.StringVar()
        self.email_var = tk.StringVar()
        title_label = tk.Label(root, text="Register New Personal Card", font=("Helvetica", 16))
        title_label.pack(pady=20)


        # Create a frame to hold labels and entries in a grid
        input_frame = tk.Frame(root)
        input_frame.pack()

        labels = ["Name:", "Stage:", "Email:"]
        for i, label_text in enumerate(labels):
            label = tk.Label(input_frame, text=label_text, width=10, anchor="e")
            label.grid(row=i, column=0, padx=(10, 5), pady=5, sticky="e")

            entry = tk.Entry(input_frame,width=50, textvariable=self.get_variable_by_index(i))
            entry.grid(row=i, column=1, padx=(0, 10), pady=5, sticky="w")

        # Buttons and canvas
        self.btn_browse_image = tk.Button(root, text="Browse Image", command=self.browse_image)
        self.btn_browse_image.pack(pady=10)

        self.label_selected_image = tk.Label(root, text="")
        self.label_selected_image.pack(pady=5)

        self.btn_generate_qr = tk.Button(root, text="Generate QR Code", command=self.generate_qr)
        self.btn_generate_qr.pack(pady=10)

        self.canvas = tk.Canvas(root, width=300, height=300)
        self.canvas.pack(pady=10)
        self.create_database()

    def create_database(self):
        self.conn = sqlite3.connect("user_data.db")
        self.cursor = self.conn.cursor()

        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                image_path TEXT,
                email TEXT,
                stage TEXT
            )
        ''')
        self.conn.commit()

    def get_variable_by_index(self, index):
        variables = [self.name_var, self.stage_var, self.email_var]
        return variables[index]
    
    def browse_image(self):
        file_path = filedialog.askopenfilename(title="Select Image", filetypes=[("Image files", "*.png;*.jpg;*.jpeg;*.gif")])
        self.label_selected_image.config(text=file_path)

    def generate_qr(self):
        name = self.name_var.get()
        image_path = self.label_selected_image.cget("text")
        email = self.email_var.get()
        stage = self.stage_var.get()


        if not name or not image_path:
            messagebox.showwarning("Error", "Please enter both name and select an image.")
            return

        user_id = self.insert_user(name, image_path, email, stage)

        if user_id is not None:
            qr_data = f"ID:{user_id},Name:{name}"

            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )

            qr.add_data(qr_data)
            qr.make(fit=True)

            img = qr.make_image(fill_color="black", back_color="white")
            img_path = f"user_qr_codes/{user_id}_qr.png"
            img.save(img_path)

            self.display_qr_code(img_path)

            self.create_word_document(user_id, name, image_path, img_path)

    def insert_user(self, name, image_path, email, stage):
        timesign = datetime.datetime.now().strftime('%H-%M-%S')
        original_img = PILImage.open(image_path)
        original_img.thumbnail((120, 120))  # Resize the image
        pth = f'storge/{timesign}_{str(name).replace(" ", "")}_image.png'
        original_img.save(pth)

        try:
            self.cursor.execute("INSERT INTO users (name, image_path, email, stage) VALUES (?, ?, ?, ?)", (name, pth, email, stage))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print("Error inserting user:", e)
            return None

    def display_qr_code(self, img_path):
        self.canvas.delete("all")

        img = Image.open(img_path)
        img = img.resize((300, 300), Image.ANTIALIAS if hasattr(Image, 'ANTIALIAS') else 3)
        img = ImageTk.PhotoImage(img)

        self.canvas.create_image(0, 0, anchor=tk.NW, image=img)
        self.canvas.image = img

    def create_word_document(self, user_id, name, image_path, qr_code_path):
        phone = self.stage_var.get()
        email = self.email_var.get()


        main_image_path = "Frame.png" 
        main_image = Image.open(main_image_path)


        image1_path = image_path
        image1 = Image.open(image1_path)


        size1 = (137, 137) 
        resized_image1 = image1.resize(size1)


        position1 = (134, 193)

        main_image.paste(resized_image1, position1, resized_image1.convert("RGBA"))

        image2_path = qr_code_path 
        image2 = Image.open(image2_path)

        size2 = (137, 137) 
        resized_image2 = image2.resize(size2)

        position2 = (764, 304) 

        main_image.paste(resized_image2, position2, resized_image2.convert("RGBA"))

        draw = ImageDraw.Draw(main_image)
        text1 = f"{name}"
        font1 = ImageFont.load_default().font_variant(size=20)
        position_text1 = (416, 196)  
        text_color1 = (0, 0, 0)  
        draw.text(position_text1, text1, font=font1, fill=text_color1)

        text2 = f"{user_id}"
        font2 = ImageFont.load_default().font_variant(size=20)
        position_text2 = (373, 239) 
        text_color2 = (0, 0, 0) 


        draw.text(position_text2, text2, font=font2, fill=text_color2)

        text3 = f"{phone}"
        font3 = ImageFont.load_default().font_variant(size=20)
        position_text3 = (405, 280) 
        text_color3 = (0, 0, 0)

        draw.text(position_text3, text3, font=font3, fill=text_color3)

        text4 = f"{email}"
        font4 = ImageFont.load_default().font_variant(size=20)
        position_text4 = (404, 320) 
        text_color4 = (0, 0, 0)

        draw.text(position_text4, text4, font=font4, fill=text_color4)

        main_image.save("presonal_image.png")
        
        main_image_path = "presonal_image.png"
        main_image = Image.open(main_image_path)
        persone_name = str(name).replace(" ", '_')
        pdf_output_path = f"{persone_name}.pdf"

        pdf_canvas = canvas.Canvas(pdf_output_path, pagesize=(main_image.width, main_image.height))

        if main_image.mode != "RGB":
            main_image = main_image.convert("RGB")

        pdf_canvas.drawImage(main_image_path, 0, 0, width=main_image.width, height=main_image.height)

        pdf_canvas.save()
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")

        # Construct the destination path on the desktop
        destination_path = os.path.join(desktop_path, pdf_output_path)

        # Move the file to the desktop
        shutil.move(pdf_output_path, destination_path)

if __name__ == "__main__":
    root = tk.Tk()
    app = UserRegistration(root)
    root.mainloop()
