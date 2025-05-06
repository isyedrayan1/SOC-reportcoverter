import streamlit as st
import os
from utils import process_excel_data
from generate_pptx import generate_pptx

# Streamlit UI
st.title("SOC Monthly Report Generator")
st.write("Upload an Excel file to generate a PowerPoint report based on the provided template.")

# File upload
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Month/Year input (optional)
month_year = st.text_input("Enter Month and Year (e.g., April 2025)", "Month X 2025")

# Generate and download button
if st.button("Generate Report"):
    if uploaded_file is not None:
        try:
            # Process Excel data
            data = process_excel_data(uploaded_file)

            # Override month_year if provided by user
            if month_year != "Month X 2025":
                data["month_year"] = month_year
                data["month"] = month_year.split(" ")[0]

            # Generate PPTX
            sanitized_month_year = month_year.replace(" ", "_")  # e.g., April_2025
            output_file = f"SOC_Monthly_Report_{sanitized_month_year}.pptx"
            generate_pptx(data, output_file)

            # Provide download link
            with open(output_file, "rb") as f:
                st.download_button(
                    label="Download PowerPoint Report",
                    data=f,
                    file_name=output_file,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            st.success("Report generated successfully!")
        except Exception as e:
            st.error(f"Error: {str(e)}")
    else:
        st.error("Please upload an Excel file.")

# Clean up generated file
if os.path.exists("SOC_Monthly_Report.pptx"):
    os.remove("SOC_Monthly_Report.pptx")