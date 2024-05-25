import qrcode
from PIL import Image

def generate_qr_code_with_image(text, image_path, output_path=r"Logo\qrcode.png"):
    # Generate the QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=0,
    )
    qr.add_data(text)
    qr.make(fit=True)
    qr_image = qr.make_image(fill_color="black", back_color="white")

    # Load the image to be placed in the center
    center_image = Image.open(image_path)

    # Resize the center image to fit in the QR code
    qr_size = qr_image.size
    center_image.thumbnail((qr_size[0] // 4, qr_size[1] // 4))

    # Calculate the position to place the center image
    center_position = ((qr_size[0] - center_image.size[0]) // 2, (qr_size[1] - center_image.size[1]) // 2)

    # Paste the center image onto the QR code
    qr_image.paste(center_image, center_position)

    # Save the final QR code image
    qr_image.save(output_path)

