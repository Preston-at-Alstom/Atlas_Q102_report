import backend
import streamlit as st

                   

# downloaded_file = 'Atlas - L102 - Delay to Train Details.xlsx'

# backend.process_report(downloaded_file)

with st.form("my_form"): 
   
   atlas_report = st.file_uploader('Upload ATLAS Q102 excel report') 
   
   st.form_submit_button('Generate Report') 

if atlas_report is not None:
   
   generated_file_names, dates_without_delays = backend.process_report(atlas_report) 
   
   for  filename in generated_file_names:
        with open(filename, 'rb') as file: 
            st.download_button(label = filename, 
                        data = file, 
                        file_name = filename, 
                        mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                        ) 
    
   if len(dates_without_delays) > 0:
       for line in dates_without_delays:
        st.write(line) 