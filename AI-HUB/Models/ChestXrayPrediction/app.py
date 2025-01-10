

import streamlit as st
from PIL import Image
import torch
from torchvision import transforms
from transformers import AutoModelForImageClassification
import pandas as pd

# Load your model
@st.cache_data
def load_dataset():
    dataset_path = "./Data_Entry_2017_v2020.csv"  # Replace with your dataset path 
    return pd.read_csv(dataset_path)

data = load_dataset()

@st.cache_resource
def load_model():
    # Define the model architecture
    model = AutoModelForImageClassification.from_pretrained("google/vit-base-patch16-224-in21k", num_labels=15)
    # Load the saved state dictionary
    state_dict = torch.load("best_model_new_retrain.pth", map_location=torch.device('cpu'))
    model.load_state_dict(state_dict)
    model.eval()
    return model

model = load_model()

# Define image transformation
transform = transforms.Compose([
    transforms.Resize((224, 224)),  # Adjust based on your model's requirements
    transforms.ToTensor(),
    transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225])  # ImageNet stats
])

# Function to make predictions
def predict_image(image):
    image = transform(image).unsqueeze(0)  # Add batch dimension
    with torch.no_grad():
        outputs = model(image).logits
        probabilities = torch.sigmoid(outputs)
    return probabilities

# Streamlit App
st.title("Chest Xray Disease Prediction App")
st.write("Upload single or multiple images to get predictions.")

# File uploader for single or bulk images
uploaded_files = st.file_uploader("Upload Image(s)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

# Process each uploaded file
if uploaded_files:
    for uploaded_file in uploaded_files:
        # Load and display the image
        image = Image.open(uploaded_file).convert("RGB")
        st.image(image, caption=f"Uploaded Image: {uploaded_file.name}", use_column_width=True)
        
        # Search for the filename in the dataset
        uploaded_filename = uploaded_file.name
        matching_row = data[data['Image Index'] == uploaded_filename]
        truth = matching_row.iloc[0]['Finding Labels'] if not matching_row.empty else "No matching label found"
        
        st.write(f"**Truth (Ground Truth Labels):** {truth}")

        # Get predictions
        probabilities = predict_image(image)

        # Create a DataFrame to display probabilities
        label_columns = [
            'No Finding', 'Infiltration', 'Effusion', 'Atelectasis', 'Nodule',
            'Mass', 'Pneumothorax', 'Consolidation', 'Pleural_Thickening',
            'Cardiomegaly', 'Emphysema', 'Edema', 'Fibrosis', 'Pneumonia', 'Hernia'
        ]
        prediction_df = pd.DataFrame({
            "Class": label_columns,
            "Probability": probabilities.squeeze().tolist()
        })

        # Highlight the highest probabilities (you can customize the threshold)
        prediction_df['Highlight'] = prediction_df['Probability'] > 0.5

        # Display predictions
        st.write("**Prediction (Model Probabilities):**")
        st.dataframe(
            prediction_df.style.format({"Probability": "{:.2f}"}).applymap(
                lambda val: 'background-color: yellow;' if val else '', subset=['Highlight']
            )
        )

