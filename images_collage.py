from docxtpl import DocxTemplate

from PIL import Image,ImageOps,ImageDraw
import os
import math
from docx.shared import Cm
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from pptx.enum.shapes import MSO_SHAPE_TYPE
def place_image_in_center():
    # Open the background image
    background_image = Image.open(r"Logo\rectangle.png")

    # Open the overlay image
    overlay_image = Image.open("collage.png")

    # Calculate the desired width and height of the overlay image
    desired_width = int(background_image.width)
    desired_height = int(background_image.height)

    # Resize the overlay image
    overlay_image = overlay_image.resize((int(desired_width//1.2), int(desired_height//1.2)))

    # Convert the overlay image to RGBA mode
    overlay_image = overlay_image.convert("RGBA")

    # Calculate the center coordinates
    x = (background_image.width - overlay_image.width) // 2
    y = (background_image.height - overlay_image.height) // 2

    # Paste the overlay image onto the background image at the center position
    background_image.paste(overlay_image, (x, y), overlay_image)

    # Save the resulting image
    background_image.save("collage.png")



def set_image(doc):
    folder_path = 'Collage'
    part_of_string = "05865-23"
    matching_files = []

    # Print the current working directory

    try:
        for filename in os.listdir(folder_path):
            if part_of_string in filename and "docx" not in filename:
                matching_files.append(os.path.join(str(os.getcwd())+'\Collage', filename))
    except OSError as e:
        print(f"Error occurred while accessing folder '{folder_path}': {e}")
    image_files = matching_files

    # Calculate the number of images and grid size
    num_images = len(image_files)
    grid_size = math.ceil(math.sqrt(num_images))

    # Calculate the maximum width and height of each image
    max_image_width = 0
    max_image_height = 0

    for image_file in image_files:
        image_path = os.path.join(folder_path, image_file)
        img = Image.open(image_path)
        width, height = img.size
        max_image_width = int(max(max_image_width, width))
        max_image_height = int(max(max_image_height, height))
    # Calculate the height of each row in the collage to connect the upper and lower images
    row_height = max_image_height

    # Calculate the width and height of the collage based on the grid size and row height
    collage_width = (max_image_width * grid_size)
    collage_height = (row_height * ((num_images - 1) // grid_size + 1))

    # Create a new blank image for the collage
    collage = Image.new('RGB', (collage_width, collage_height), 'white')

    # Iterate over the image files and resize/rotate them to fit the collage
    x_offset = 0
    y_offset = 0

    for image_file in image_files:
        image_path = os.path.join(folder_path, image_file)
        img = Image.open(image_path)

        # Check if the image height is greater than the width
        width, height = img.size
        if height > width:
            img = img.rotate(90, expand=True)

        # Resize the image to fit the row height
        img = img.resize((max_image_width, row_height))

        x = x_offset * max_image_width
        y = y_offset * row_height

        collage.paste(img, (x, y))

        x_offset += 1
        if x_offset >= grid_size:
            x_offset = 0
            y_offset += 1
    collage.save('collage.png')
    place_image_in_center()


