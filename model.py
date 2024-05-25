from PIL import Image, ImageDraw, ImageFont
def add_arrow(doc):
    table_index = 4  # Index starts from 0
    table = doc.tables[table_index]
    img = Image.open(r'static\arrow.png')
    for row in table.rows:
        cell_text = row.cells[2].text  # Assuming the third column is column index 2
        if cell_text.strip():  # Check if cell text is not empty
            label_name = cell_text.strip()  # Remove extra spaces
            image_width, image_height = img.size
            transparent_section_width = 200  # Adjust this value as needed

            # Create a new image with an RGBA mode (includes alpha channel for transparency)
            new_img = Image.new('RGBA', (image_width + transparent_section_width, image_height), (0, 0, 0, 0))

            # Paste the original image onto the new image with an offset
            new_img.paste(img, (transparent_section_width, 0))

            # Call draw Method to add 2D graphics to the new image
            draw = ImageDraw.Draw(new_img)

            # Custom font style and font size
            font_path = r'static\cambria.ttf'
            font_size = 40
            myFont = ImageFont.truetype(font_path, font_size)

            # Add colored label to the transparent section
            label_text = label_name
            label_position = (20, 10)  # Adjust the position as needed
            label_color = (232, 93, 68)  # RGB color for #e85d44
            draw.text(label_position, label_text, font=myFont, fill=label_color)
            edited_image_path = fr"static\temp_images\{label_name}.png"
            new_img.save(edited_image_path)

            # Optional: Print a message once the image is saved
            print(f"Edited image saved as {edited_image_path}")
