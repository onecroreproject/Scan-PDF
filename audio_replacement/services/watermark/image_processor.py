import os

from PIL import Image


class WatermarkImageProcessor:
    """
    Resize watermark and apply opacity.
    """

    def __init__(
        self,
        image_path,
    ):

        self.image_path = image_path

    def process(
        self,
        scale=20,
        opacity=70,
    ):

        image = Image.open(
            self.image_path
        ).convert("RGBA")

        width, height = image.size

        new_width = max(
            1,
            int(width * scale / 100)
        )

        ratio = new_width / width

        new_height = int(
            height * ratio
        )

        image = image.resize(
            (
                new_width,
                new_height,
            ),
            Image.LANCZOS,
        )

        alpha = image.getchannel("A")

        alpha = alpha.point(
            lambda p: int(
                p * opacity / 100
            )
        )

        image.putalpha(alpha)

        output_path = os.path.splitext(
            self.image_path
        )[0] + "_processed.png"

        image.save(
            output_path,
            "PNG",
        )

        return output_path