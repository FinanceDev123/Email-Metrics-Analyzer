# Email-Metrics-Tool
Description:

A Python-based tool to search, filter, and analyze emails using keywords and date ranges. Includes functionality to retrieve email metadata, count keyword occurrences, and display insights.



Enabling Your Email for Python Access:
1. Enable IMAP Access
• Log in to your email account (e.g., Gmail).
• Go to Settings > See all settings > Forwarding and POP/IMAP.
• Under the IMAP access section, select Enable IMAP.
• Click Save Changes.

2. Allow Less Secure App Access (if required)
Some email providers, such as Gmail, require you to enable less secure app access:
• Visit Google Account Security.
• Scroll down to Less secure app access and toggle it ON.
(Note: Google may block this option in the future. Use App Passwords instead if available.)

3. Use an App Password (recommended for Gmail and other secure services)
• Go to your email account’s security settings and enable 2-Step Verification.
• Once enabled, generate an App Password:
  • Under the Security tab, locate App Passwords.
  • Select the app and device (e.g., Python script on Windows).
  • Generate the password and use it instead of your normal password.

4. Additional Provider Notes
• For Gmail users, IMAP Server: imap.gmail.com, Port: 993.
• For Outlook users, IMAP Server: outlook.office365.com, Port: 993.

Important: Be cautious when using your email credentials in scripts. Avoid sharing or exposing them. Use environment variables or encrypted configuration files whenever possible. Alternatively, if you only want to expirement with the code you can always create a throwaway email account.

Thank you for your attention.🙏

   
