import io
import base64
from pathlib import Path

from PIL import Image


def get_element_text(element):
    """get all text of an element
    """
    try:
        children = element.xpath('.//w:t')  # not working for lxml.etree._Element
    except:
        children = element.iterchildren()
    return ''.join(c.text for c in children if c.text).strip()


def blob_to_image(blob,
                  image_as='base64',
                  image_type='jpeg',
                  filename='image',
                  media_dir=Path('.'),
                  ):
    """convet image blob data to image file or base64 string
    """
    image = Image.open(io.BytesIO(blob))

    if image_type.lower() in ('jpeg', 'jpg'):
        image_type = 'jpeg'
        image = image.convert('RGB')               # png => jpeg, smaller size
        filename = f'{filename}.jpg'
    else:
        filename = f'{filename}.png'

    if image_as == 'file':
        if not media_dir.exists():
            media_dir.mkdir(parents=True)
        image.save(media_dir.joinpath(filename))
        image = str(media_dir.joinpath(filename))
    else:
        buffered = io.BytesIO()
        image.save(buffered, image_type)
        prefix = f'data:image/{image_type};base64,'.encode()
        image = prefix + base64.b64encode(buffered.getvalue())
        image = image.decode()

    return image, filename
