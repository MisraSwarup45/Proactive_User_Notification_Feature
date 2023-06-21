### Proactive_User_Notification_Feature
### LinkedIn Unread Messages and Notifications Monitor

This script monitors the number of unread messages and notifications on LinkedIn and sends an email with the updated counts. It utilizes Selenium WebDriver and OpenPyXL to retrieve the data and store it in an Excel file.

## Requirements

- Python 3.0 or above
- ChromeDriver
- Selenium WebDriver
- OpenPyXL

## Installation

1. Clone the repository or download the script file.

2. Install the required Python packages using pip:

   ```bash
   pip install selenium openpyxl
   ```
3. Download ChromeDriver and place it in the directory of the script. Make sure to match the ChromeDriver version with your Chrome browser version.

4. Configure the following parameters in the script:

   - `username`: Your LinkedIn username
   - `password`: Your LinkedIn password
   - `chromedriver_path`: Path to the ChromeDriver executable
   - `sender_email`: Your email address
   - `sender_password`: Your email password or app password
   - `recipient_email`: Email address to receive the notifications
   - `smtp_server`: SMTP server address
   - `smtp_port`: SMTP server port

. Run the script:

   ```bash
   python Linkedin.py
   ```
## Acknowledgements

- Selenium WebDriver
- OpenPyXL
