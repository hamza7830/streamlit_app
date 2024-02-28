import streamlit as st
import pandas as pd
from pptx import Presentation  
from Excel_to_Ppt import process_files

if __name__ == "__main__":
    
    st.title('PPT File Generator')
    adscore_file = st.file_uploader("Upload the Adscore Excel file")
    example_excel_url = st.file_uploader("Enter the Example Excel file URL")
    ppt_file = st.file_uploader("Upload the original PowerPoint file", type=['pptx'])
    st.title('Choose the Slide Option')
    options = ['Airline', 'Travel', 'Finance', 'Hotel','Auto','Tech','FDI','Others']
    selected_option = st.selectbox('Which option do you think is the best?', options)
    if st.button('Process Files'):
        if adscore_file and example_excel_url and ppt_file:
            updated_ppt = process_files(adscore_file, example_excel_url, ppt_file,selected_option)
            print("++++++++++++++++++++++++++++++++++   E   N   D   +++++++++++++++++++++++++++++++++++++++++++++++r")
        else:
            st.error("Please upload all files and provide the necessary URLs to proceed.")
