SOC Monthly Report Generator
Overview
This tool automates the generation of a PowerPoint (PPTX) report for a SOC monthly report using data from an Excel file. It replicates the structure of a predefined template (21 slides) and includes a Streamlit UI for easy interaction.
Setup

Clone the repository:git clone <repository-url>
cd <repository-directory>

If you have downloaded zip file:

1.Unzip and open terminal and go to project directory and make an virtual env if there isnt any venv by typing : python -m venv .venv

2.If there is venv created or already have activate it by typing: .venv\Scripts\activate

3.Install dependencies: pip install -r requirements.txt

Note: Always download the dependencies in virtual environment(venv) by creating(step1) and activating it(step2)

Now to run the app run this command
Run the Streamlit app: streamlit run app.py



Usage

Open the Streamlit app in your browser (default: http://localhost:8501).
Upload an Excel file containing SOC data (expected columns: Case Number, Date/Time, Subject, Priority, Disposition, Report Status, Alert Type, etc.).
Optionally, enter the month and year for the report (e.g., "April 2025"). If not provided, the tool will attempt to extract it from the Excel or use a placeholder.
Click "Generate Report" to create the PPTX.
Download the generated PowerPoint using the provided button.

File Structure

app.py: Streamlit UI script.
generate_pptx.py: Logic for generating the PPTX.
config.py: Slide configuration.
utils.py: Helper functions for data processing and PPTX creation.
requirements.txt: Project dependencies.

Notes

The Excel file must match the expected structure (see prior analysis for details).
The tool assumes previous month data for comparisons (placeholders are used if unavailable).
Placeholders are used for corrupted slides (12–14) and missing data.

