from PIL import Image
from io import BytesIO
import pytesseract


pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"


class CaptchaHandler:
    pixel_range = [(i, i, i) for i in range(102, 130)]

    def get_captcha_from_image(self, target_image):
        image_data = BytesIO(target_image)

        image = Image.open(image_data)
        width, height = image.size

        image = image.convert("RGB")

        white_image = Image.new("RGB", (width, height), "white")

        for x in range(width):
            for y in range(height):
                pixel = image.getpixel((x, y))

                if pixel in self.pixel_range:
                    white_image.putpixel((x, y), pixel)

        return pytesseract.image_to_string(white_image).strip()