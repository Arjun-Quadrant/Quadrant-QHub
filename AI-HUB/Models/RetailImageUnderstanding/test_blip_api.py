import base64
import io

import requests
from PIL import Image

API_URL = "https://xt41z6d0qly7bg8w.us-east-1.aws.endpoints.huggingface.cloud"
headers = {
    "Accept": "application/json",
    "Content-Type": "application/json"
}


def query(payload):
    response = requests.post(API_URL, headers=headers, json=payload)
    return response.json()


image_path = 'data/train/images/111736998195_1.JPG'  # Replace with the path to your image file

# Open the image and convert it to bytes
with Image.open(image_path) as img:
    img_bytes = io.BytesIO()
    img.save(img_bytes, format='JPEG')
    image_bytes = img_bytes.getvalue()
    image_base64 = base64.b64encode(image_bytes).decode('utf-8')

output = query({
    "inputs": {
        "text": "",
        "image": image_base64
    },
    "parameters": {
        "max_new_tokens": 150
    }
})

print(output)