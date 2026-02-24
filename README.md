Dynamic-reporting-system-excel

A comprehensive end-to-end Excel data analytics project.
Includes automated data cleaning via Power Query, complex data modeling, and a dynamic, 
interactive dashboard utilizing Pivot Tables and advanced visualization techniques to provide actionable business insights.
Project Overview:-
This project eliminates the manual effort of data entry and report formatting. 
It provides a Python-powered desktop application that transforms raw, unorganized CSV files into structured, multi-sheet Excel workbooks.
The system automates data aggregation, calculates summary statistics, and embeds visual charts into the final output.

Key Features:-
One-Click Automation: A Tkinter-based GUI allows users to simply upload a CSV and receive a fully formatted Excel file.
Automated Summarization: Uses Pandas to generate Pivot Tables, aggregating metrics like total sales or salaries per department.
Statistical Insights: Automatically calculates key descriptive statistics (Mean, Median, Standard Deviation) for the dataset.
Visual Integration: Generates dynamic bar charts using Matplotlib and embeds them directly into specific Excel cells using the Openpyxl engine.
Multi-Sheet Architecture: Organizes output into "Raw Data," "Summary Report," and "Statistics" for professional presentation.

Tools Used:-
Python: Core programming language.
Pandas: For data cleaning and pivot table logic.
Matplotlib: For generating data visualizations.
Openpyxl: For Excel file styling and image embedding.
Tkinter: For the graphical user interface and file dialogs.

Project Structure:-
/main.py: The primary application script containing the GUI and processing logic.
/data.csv: Sample dataset used for testing the reporting pipeline.
/Final_Report.xlsx: Example of the generated output with charts and summaries.
/report_chart.png: Temporary image file generated during the visualization phase.

How to Use:-
Clone the repo and ensure you have the libraries installed: pip install pandas openpyxl matplotlib.
Run the script: Execute python main.py.
Upload Data: Use the "Upload CSV" button to select your raw data file.
Save Report: Choose your destination and filename. The tool will process the data and notify you upon completion.

Insights Derived:-
Operational Efficiency: Automated a manual reporting task, reducing processing time from 30 minutes to under 3 seconds.
Departmental Trends: Isolated high-spending departments through automated pivot table grouping.
Visual Clarity: Provided instant trend identification via embedded bar charts, making raw numbers easier to interpret.

                           <img width="512" height="183" alt="image" src="https://github.com/user-attachments/assets/45efdfc4-167e-4651-ac82-0944cf55eeac" />

                           <img width="862" height="380" alt="Excel_output" src="https://github.com/user-attachments/assets/e3e54c9b-273d-4d9b-a8d4-471ddee94a7f" />

