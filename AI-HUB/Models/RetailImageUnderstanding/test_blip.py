import requests
from PIL import Image
from transformers import BlipProcessor, BlipForConditionalGeneration
from utils import image_utils

processor = BlipProcessor.from_pretrained("quadranttechnologies/qhub-blip-image-captioning-finetuned")
model = BlipForConditionalGeneration.from_pretrained("quadranttechnologies/qhub-blip-image-captioning-finetuned")

img_url = 'https://storage.googleapis.com/sfr-vision-language-research/BLIP/demo.jpg'
img_path = 'data/train/images/111736998195_1.JPG'
raw_image = image_utils.load_image(img_url)

# conditional image captioning
text = ""
inputs = processor(raw_image, text, return_tensors="pt")

out = model.generate(**inputs)
print(processor.decode(out[0], skip_special_tokens=True))
# >>> a photography of a woman and her dog

# unconditional image captioning
inputs = processor(raw_image, return_tensors="pt")

out = model.generate(**inputs)
print(processor.decode(out[0], skip_special_tokens=True))

