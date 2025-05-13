# LinkedIn Referral Assistant

A Python tool to automatically process LinkedIn referral requests, extract resumes, and compile key information into an Excel sheet.

## Overview

This tool helps you manage LinkedIn job referral requests by:

1. Reading all unread LinkedIn messages
2. Using AI to identify which messages are referral requests
3. Extracting resumes (from attachments or Google Drive links)
4. Saving resumes in a local folder
5. Extracting key information like name, email, phone, and years of experience
6. Compiling all information into an Excel sheet

## Requirements

- Python 3.7+
- Google Chrome browser with a logged-in LinkedIn account
- Azure OpenAI API access (for GPT analysis)

## Installation

1. Clone this repository:
```
git clone https://github.com/yourusername/linkedin-referral-assistant.git
cd linkedin-referral-assistant
```

2. Install the required dependencies:
```
pip install -r requirements.txt
```

3. Install the Playwright browsers:
```
python -m playwright install
```

## Configuration

Before running the script, make sure to update your Azure OpenAI credentials in the `analyze_referrals.py` file:

```python
# Azure OpenAI Configuration
client = AzureOpenAI(
    api_key="YOUR_API_KEY",
    azure_endpoint="YOUR_AZURE_ENDPOINT",
    api_version="2024-02-15-preview"
)
```

## Usage

Run the main script:
```
python linkedin_referral_assistant.py
```

The script will:
1. Prompt you for your Chrome profile path (press Enter to use the default)
2. Open a browser window to access your LinkedIn messages
3. Process unread messages and analyze them for referral requests
4. Save resumes to the `resumes/` folder
5. Create an Excel file `referral_requests.xlsx` with the compiled information

## Output

- **Resumes folder**: Contains all extracted PDF resumes, named after the sender
- **Excel file**: Contains columns for:
  - Job ID (if provided in the message)
  - Name
  - Email
  - Phone number
  - Years of work experience
  - Resume path (local file location)
  - Position applied for
  - Company

## Notes

- The script uses your existing Chrome profile to access LinkedIn, so you don't need to log in again
- It only processes unread messages, so mark messages as read if you don't want them to be processed again
- For Google Drive links, the script attempts to download the resume directly

## Troubleshooting

- If you encounter a "Profile not found" error, make sure to provide the correct Chrome profile path
- If LinkedIn's UI changes, the selectors in the code might need to be updated
- Make sure your LinkedIn account is logged in within your Chrome profile

## License

This project is licensed under the MIT License - see the LICENSE file for details.
