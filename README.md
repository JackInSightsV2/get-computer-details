# ComputerDetails.vbs

This VBScript collects and displays detailed computer information in an HTML page using Windows Management Instrumentation (WMI). It gathers system, processor, operating system, network, graphics card, and monitor details. The generated HTML page is styled with a monospaced font and includes three interactive buttons:

- **Copy to Clipboard:** Copies the page’s text to the clipboard.
- **Email:** Opens your default mail client with the page text prefilled in the email body.
- **Save as Text File:** Saves the page text as a text file ("ComputerDetails.txt").

## Features

- Collects comprehensive computer details via WMI.
- Displays information such as Computer Name, User Name, Manufacturer, Model, Memory, Processor, OS details, and enhanced Network Details (including Ethernet Adapter Name, IP Address, Subnet, and Default Gateway).
- Formats the output using HTML with left-aligned, bold labels and unbolded values.
- Uses Courier New (monospace) for consistent formatting.
- Provides three interactive buttons at the bottom of the page for additional functionality.

## Usage

1. Ensure you are running on a Windows operating system with WScript support.
2. Double-click the `ComputerDetails.vbs` file or run it from the command line.
3. The script will gather system details, generate an HTML file in your temporary folder, and open it in your default web browser.
4. Use the buttons at the bottom of the page to copy the content, email it, or save it as a text file.

## Requirements

- Windows Operating System

## License

MIT License

## Author

Stephen Henry  
GitHub: https://github.com/JackInSightsV2/
