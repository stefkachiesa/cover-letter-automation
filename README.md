# Cover Letter Automation Project

This project automates the creation, customization, and sharing of professional cover letters. It uses Python libraries to generate cover letters based on a template, converts them to PDF format, and uploads them to Google Drive for easy sharing.

## Disclaimer

The cover letter that includes a link to this repository may have been created with a slightly different version of the code, a different template, or other variations.


## Features

- **Customizable Templates**: Easily generate cover letters with dynamic placeholders for position, company name, sector, achievements, and strengths.
- **Professional Formatting**: Ensures consistent and professional formatting using the `python-docx` library.
- **Automated PDF Conversion**: Converts DOCX files to PDF format for universal readability.
- **Seamless Cloud Integration**: Uploads the generated PDF to Google Drive and sets appropriate permissions for easy sharing.
- **Robust Error Handling**: Includes comprehensive error handling to ensure reliability.

## Prerequisites

- Python 3.x
- `python-docx` library
- `docx2pdf` library
- `google-api-python-client` library
- `google-auth` library
- A Google Cloud project with Google Drive API enabled
- Service account credentials JSON file

## Installation

 ## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/stefkachiesa/cover-letter-automation.git
    cd cover-letter-automation
    ```

2. Install the required libraries:
    ```sh
    pip install python-docx docx2pdf google-api-python-client google-auth
    ```

3. Place your service account credentials JSON file in the project directory and rename it.

## Usage

To create a cover letter, write your own template, customize the parameters as needed and run the script:

```python
create_cover_letter_your_name(
    position="Data Analyst",
    company_name="TechCorp",
    sector="technology",
    achievements="innovative projects",
    strengths="analytical skills and problem-solving abilities",
    font_name="Arial",
    font_size=12
)


