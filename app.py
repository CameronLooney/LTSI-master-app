import os
import streamlit as st
import numpy as np
from PIL import  Image

# Custom imports
from multipage import MultiPage
from pages import email, feedback, open_orders # import your pages here

# Create an instance of the app
app = MultiPage()

# Title of the main page
display = Image.open('logo.png')

st.image(display, width = 650)
# st.title("Data Storyteller Application")



# Add all your application here
app.add_page("Generate Open Order File", open_orders.app)
app.add_page("Email C-SAMS", email.app)
app.add_page("Consolidate Feedback", feedback.app)


# The main app
app.run()