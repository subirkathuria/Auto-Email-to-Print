\# Auto Email-to-Print Script 🖨️



A lightweight, secure PowerShell script that automatically scans a folder for files, emails them to a specific address (like an HP ePrint or Epson Connect printer), and archives the sent files.



\## Features

\- \*\*Zero-Touch Automation:\*\* Runs silently in the background.

\- \*\*Printer Friendly:\*\* Sends emails with blank subjects and bodies so printers only print the attachment.

\- \*\*Secure:\*\* Encrypts your Gmail App Password using Windows DPAPI (standard Windows encryption).

\- \*\*Smart Archiving:\*\* Automatically moves processed files to a `Sent\_Files` folder, handling duplicate filenames automatically.

\- \*\*Batching:\*\* Sends files in batches of 10 to avoid email server limits.



\## Prerequisites

\- Windows 10 or 11

\- A Gmail account with \*\*2-Step Verification\*\* enabled.

\- An \*\*App Password\*\* from Google (see \[Google's Guide](https://support.google.com/accounts/answer/185833)).



\## Installation



1\. Download the `Send-Files.ps1` script.

2\. Right-click the file, go to \*\*Properties\*\*, and check \*\*Unblock\*\* (if visible).



\## Usage



\*\*1. First Run (Setup)\*\*

Double-click the script (or run via PowerShell). It will prompt you for:

\- Your Gmail address.

\- Your 16-character App Password.

\- The destination email (e.g., `printer@hpeprint.com`).

\- The folder you want to monitor.



\*\*2. Automatic Mode\*\*

Create a shortcut to the script and add the following Target path to make it run silently:

`powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File "C:\\Path\\To\\Send-Files.ps1"`



\*\*3. Resetting Credentials\*\*

To clear saved passwords and settings, run the script with the reset flag:

`.\\Send-Files.ps1 --reset`



\## Security Note

This script stores your credentials locally on your machine using standard Windows encryption. The password file (`mailer\_secure\_string.txt`) can only be decrypted by the logged-in user on that specific PC.



\## License

MIT License - feel free to modify and use!

