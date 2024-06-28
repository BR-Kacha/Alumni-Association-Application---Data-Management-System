# Alumni-Association-Application---Data-Management-System

## Overview
This application is designed to efficiently manage the alumni data of Atmiya University by organizing it into various departments, branches, and graduation years. It creates separate Excel files for each branch within their respective department directories, removes duplicate entries, and provides visual representations of the yearwise data. This tool saves significant time compared to manual data sorting, providing quick and accurate results in just a few seconds.

## Author
- Brijraj R. Kacha | LinkedIn: [https://www.linkedin.com/in/brijraj-kacha/]

## Features
- Filters and organizes alumni data by departments, branches, and graduation years.
- Creates and updates Excel files for each branch in the respective department’s directory.
- Removes duplicate entries.
- Provides visual representations of yearwise data in the form of graphs.
- Automatically generates required directories and files.
- Easy-to-use GUI.
- Available as a standalone .exe file for easy execution.

## Usage
### Running the Executable (.exe)
1. Download the `alumni_data_management.exe` file from the repository.
2. Execute the file by double-clicking on it.
3. Upload the input data file when prompted.

### Input Data Format
Ensure that the input data file is in the same format as the provided `input_data_format.xlsx` file. The input file should contain the necessary columns for organizing the data correctly.

### Output
The application will generate the following:
- Excel files for each branch within their respective department directories.
- Visual representations of the year data in the form of graph.

### Directory Structure
- The directories will be created automatically when the application is executed.
- If the directories already exist, they will be updated with the new data.
- No need to manually create directories or Excel files.

### Example
An example of the input data format and the output directory structure is provided in the repository. Note that the actual data has been removed for privacy reasons, and only the empty files with row and column names are shared.

## Files in the Repository
- `app.py`: The final GUI Python code for the application.
- `alumni_data_management.exe`: The executable file for easy execution.
- `input_data_format.xlsx`: The input data format file.
- `sample_output`: A directory containing the example output structure with empty files.
- `SRS Document.pdf`: The Software Requirement Specification Document for the project.

