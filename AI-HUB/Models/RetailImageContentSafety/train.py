from datetime import datetime

import torch
from torch import nn
from torch.utils.data import DataLoader
from transformers import CLIPModel, CLIPProcessor
from utils.image_text_datset import ImageTextDataset

def collate_fn(batch):
    images, captions = zip(*batch)
    inputs = processor(
        text=list(captions),
        images=list(images),
        return_tensors="pt",
        padding=True,  # Ensures that text inputs are padded to the same length
        truncation=True,
        max_length=77
    )
    return inputs

model_name = 'openai/clip-vit-base-patch32'
model = CLIPModel.from_pretrained(model_name)
processor = CLIPProcessor.from_pretrained(model_name)

dataset = ImageTextDataset(
    annotations_file='data/train/annotations.tsv',
    img_dir='data/train/images',
    processor=processor
)

dataloader = DataLoader(dataset, batch_size=16, shuffle=True, collate_fn=collate_fn)

device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
if torch.mps.is_available():
    device = "mps"
model.to(device)

optimizer = torch.optim.AdamW(model.parameters(), lr=5e-5)
criterion = nn.CrossEntropyLoss()

epochs = 5

for epoch in range(epochs):
    model.train()
    total_loss = 0

    for batch in dataloader:
        optimizer.zero_grad()

        input_ids = batch['input_ids'].to(device)
        attention_mask = batch['attention_mask'].to(device)
        pixel_values = batch['pixel_values'].to(device)

        outputs = model(
            input_ids=input_ids,
            attention_mask=attention_mask,
            pixel_values=pixel_values,
            return_loss=True
        )

        loss = outputs.loss
        loss.backward()
        optimizer.step()

        total_loss += loss.item()

    avg_loss = total_loss / len(dataloader)
    print(f"Epoch {epoch + 1}/{epochs}, Loss: {avg_loss:.4f}")

    model.config.fine_tuned_on = "Custom annotated based on Stanford Online Products Dataset"
    model.config.fine_tuning_task = "Zero-Shot Image Classification for Content Safety"
    model.config.fine_tuned_by = "Quadrant Technologies"
    model.config.date_fine_tuned = str(datetime.now())

    model.save_pretrained('quadranttechnologies/retail-content-safety-finetuned_clip')
    processor.save_pretrained('quadranttechnologies/retail-content-safety-finetuned_clip_processor')


