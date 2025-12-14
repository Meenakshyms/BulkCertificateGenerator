import os
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageTk
from tkinter import Tk, Label, Button, filedialog, colorchooser, Canvas, Frame, messagebox, simpledialog
from datetime import datetime

# --------------------------------------------
# MAIN WINDOW
# --------------------------------------------
root = Tk()
root.title("Premium Certificate Generator")
root.geometry("1300x900")
root.configure(bg="#1c1c1c")

# --------------------------------------------
# GLOBAL VARIABLES
# --------------------------------------------
data_file = ""
save_folder = ""
bg_texture_path = None
logo_path = None

content_template = "has successfully completed the [course] by [organization] at [venue] on [date]."
text_color = "#000000"

title_font_path = "arial.ttf"
content_font_path = "arial.ttf"
signature_font_path = "arial.ttf"

current_drag = None     # which element is being dragged
offset_x = 0
offset_y = 0

# All draggable item positions
positions = {
    "title": (400, 60),
    "name": (400, 150),
    "content": (400, 250),
    "signature_label": (150, 380),
    "signature_line": (150, 370),
    "logo": (680, 40)
}

# Stores custom text boxes: {id: {"text":..., "pos":(x,y), "font":..., "color":...}}
custom_text_items = {}
custom_id_counter = 0

# --------------------------------------------
# FILE SELECTION
# --------------------------------------------
def select_data_file():
    global data_file
    path = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
    if path:
        data_file = path
        data_label.config(text=f"Data File: {os.path.basename(path)}")
        preview_certificate()

def select_save_folder():
    global save_folder
    save_folder = filedialog.askdirectory()
    if save_folder:
        save_label.config(text=f"Save Folder: {save_folder}")

def select_bg_texture():
    global bg_texture_path
    path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
    if path:
        bg_texture_path = path
        preview_certificate()

def select_text_color():
    global text_color
    color = colorchooser.askcolor()[1]
    if color:
        text_color = color
        preview_certificate()

def select_logo():
    global logo_path
    path = filedialog.askopenfilename(filetypes=[("Images", "*.png *.jpg *.jpeg")])
    if path:
        logo_path = path
        preview_certificate()

# --------------------------------------------
# BACKGROUND LOADING
# --------------------------------------------
def create_textured_background(width, height):
    if bg_texture_path:
        try:
            img = Image.open(bg_texture_path).convert("RGB")
            return img.resize((width, height))
        except:
            return Image.new("RGB", (width, height), "white")
    else:
        return Image.new("RGB", (width, height), "white")

# --------------------------------------------
# DRAGGABLE HANDLER
# --------------------------------------------
def on_drag_start(event):
    global current_drag, offset_x, offset_y

    clicked = preview_canvas.find_closest(event.x, event.y)[0]
    current_drag = clicked

    x0, y0, x1, y1 = preview_canvas.bbox(clicked)
    offset_x = event.x - x0
    offset_y = event.y - y0

def on_drag_move(event):
    if current_drag:
        new_x = event.x - offset_x
        new_y = event.y - offset_y
        preview_canvas.coords(current_drag, new_x, new_y)

def on_drag_end(event):
    global current_drag

    if not current_drag:
        return

    # Save updated positions
    x0, y0, _, _ = preview_canvas.bbox(current_drag)

    # Identify which item is moved
    for key, item_id in canvas_items.items():
        if item_id == current_drag:
            positions[key] = (x0, y0)

    # Custom text items
    for cid, item in custom_text_items.items():
        if item["canvas_id"] == current_drag:
            item["pos"] = (x0, y0)

    current_drag = None

# --------------------------------------------
# ADDING CUSTOM TEXT BOX
# --------------------------------------------
def add_custom_textbox():
    global custom_id_counter
    text = simpledialog.askstring("Custom Text", "Enter text:")
    if not text:
        return

    custom_id = f"custom_{custom_id_counter}"
    custom_id_counter += 1

    item = preview_canvas.create_text(
        200, 200,
        text=text,
        font=(content_font_path, 22),
        fill=text_color,
        anchor="nw"
    )

    custom_text_items[custom_id] = {
        "text": text,
        "pos": (200, 200),
        "font": (content_font_path, 22),
        "color": text_color,
        "canvas_id": item
    }

# --------------------------------------------
# PREVIEW CERTIFICATE
# --------------------------------------------
canvas_items = {}

def preview_certificate():
    preview_canvas.delete("all")

    width, height = 800, 450
    cert = create_textured_background(width, height)
    draw = ImageDraw.Draw(cert)

    # Fonts
    try:
        title_font = ImageFont.truetype(title_font_path, 40)
        content_font = ImageFont.truetype(content_font_path, 24)
        signature_font = ImageFont.truetype(signature_font_path, 24)
    except:
        title_font = ImageFont.load_default()
        content_font = ImageFont.load_default()
        signature_font = ImageFont.load_default()

    # Convert PIL image to Tk
    tk_img = ImageTk.PhotoImage(cert)
    preview_canvas.image = tk_img
    preview_canvas.create_image(0, 0, anchor="nw", image=tk_img)

    # Draggable elements
    canvas_items.clear()

    # Title
    canvas_items["title"] = preview_canvas.create_text(
        *positions["title"], text="Certificate of Completion",
        font=(title_font_path, 32), fill=text_color, anchor="nw"
    )

    # Name placeholder
    canvas_items["name"] = preview_canvas.create_text(
        *positions["name"], text="[Name]",
        font=(title_font_path, 30), fill=text_color, anchor="nw"
    )

    # Content
    content_preview = content_template.replace("[Name]", "[Name]").replace("[Date]", datetime.today().strftime("%Y-%m-%d"))
    canvas_items["content"] = preview_canvas.create_text(
        *positions["content"], text=content_preview,
        font=(content_font_path, 20), fill=text_color, anchor="nw"
    )

    # Signature line
    canvas_items["signature_line"] = preview_canvas.create_line(
        positions["signature_line"][0], positions["signature_line"][1],
        positions["signature_line"][0] + 200, positions["signature_line"][1],
        fill=text_color, width=2
    )

    # Signature label
    canvas_items["signature_label"] = preview_canvas.create_text(
        *positions["signature_label"], text="Instructor Signature",
        font=(signature_font_path, 20), fill=text_color, anchor="nw"
    )

    # Logo
    if logo_path:
        logo = Image.open(logo_path).resize((100, 100))
        tk_logo = ImageTk.PhotoImage(logo)
        preview_canvas.logo = tk_logo

        canvas_items["logo"] = preview_canvas.create_image(
            *positions["logo"], anchor="nw", image=tk_logo
        )

    # Restore custom textboxes
    for cid, item in custom_text_items.items():
        item["canvas_id"] = preview_canvas.create_text(
            *item["pos"], text=item["text"], font=item["font"],
            fill=item["color"], anchor="nw"
        )

# --------------------------------------------
# GENERATE CERTIFICATES
# --------------------------------------------
def generate_certificates():
    if not data_file or not save_folder:
        messagebox.showerror("Error", "Select data file and save folder.")
        return

    try:
        data = pd.read_csv(data_file) if data_file.endswith(".csv") else pd.read_excel(data_file)
    except Exception as e:
        messagebox.showerror("Error", f"Cannot read data:\n{e}")
        return

    for _, row in data.iterrows():
        width, height = 1600, 900
        cert = create_textured_background(width, height)
        draw = ImageDraw.Draw(cert)

        # Fonts
        try:
            title_font = ImageFont.truetype(title_font_path, 60)
            content_font = ImageFont.truetype(content_font_path, 35)
            signature_font = ImageFont.truetype(signature_font_path, 30)
        except:
            title_font = content_font = signature_font = ImageFont.load_default()

        # Title
        draw.text(positions["title"], "Certificate of Completion", font=title_font, fill=text_color)

        # Name
        name_text = row["Name"]
        draw.text(positions["name"], name_text, font=title_font, fill=text_color)

        # Content text
        content_text = content_template.replace("[Name]", name_text).replace(
            "[Date]", datetime.today().strftime("%Y-%m-%d")
        )
        draw.text(positions["content"], content_text, font=content_font, fill=text_color)

        # Signature line
        x, y = positions["signature_line"]
        draw.line((x, y, x + 300, y), fill=text_color, width=3)

        # Signature label
        draw.text(positions["signature_label"], "Instructor Signature", font=signature_font, fill=text_color)

        # Logo
        if logo_path:
            try:
                logo = Image.open(logo_path).resize((150, 150))
                cert.paste(logo, positions["logo"], logo if logo.mode == "RGBA" else None)
            except:
                pass

        # Custom textboxes
        for cid, item in custom_text_items.items():
            draw.text(item["pos"], item["text"], font=ImageFont.truetype(content_font_path, 30), fill=item["color"])

        # Save
        cert.save(os.path.join(save_folder, f"{name_text}_certificate.png"))

    messagebox.showinfo("Success", "Certificates generated successfully!")

# --------------------------------------------
# GUI ELEMENTS
# --------------------------------------------
Label(root, text="Premium Certificate Generator", font=("Arial", 24), bg="#1c1c1c", fg="white").pack()

frame_files = Frame(root, bg="#1c1c1c")
frame_files.pack(pady=5)

Button(frame_files, text="Select Data File", command=select_data_file).grid(row=0, column=0)
data_label = Label(frame_files, text="Data File: None", bg="#1c1c1c", fg="white")
data_label.grid(row=0, column=1)

Button(frame_files, text="Select Save Folder", command=select_save_folder).grid(row=1, column=0)
save_label = Label(frame_files, text="Save Folder: None", bg="#1c1c1c", fg="white")
save_label.grid(row=1, column=1)

frame_opt = Frame(root, bg="#1c1c1c")
frame_opt.pack(pady=5)

Button(frame_opt, text="Background Texture", command=select_bg_texture).grid(row=0, column=0)
Button(frame_opt, text="Text Color", command=select_text_color).grid(row=0, column=1)
Button(frame_opt, text="Add Logo", command=select_logo).grid(row=0, column=2)
Button(frame_opt, text="Add Custom Text Box", command=add_custom_textbox).grid(row=0, column=3)

preview_canvas = Canvas(root, width=800, height=450, bg="grey")
preview_canvas.pack(pady=10)

preview_canvas.bind("<Button-1>", on_drag_start)
preview_canvas.bind("<B1-Motion>", on_drag_move)
preview_canvas.bind("<ButtonRelease-1>", on_drag_end)

Button(root, text="Generate Certificates", font=("Arial", 16), bg="#00ffaa",
       command=generate_certificates).pack(pady=20)

preview_certificate()
root.mainloop()