# Anti-Spoof Phishing

Program that monitors Outlook inbox in real-time to identify spoof phishing emails sent out by third-party security consultants. Places them in the junk email where they belong.

## Getting Started

- This program requires Python whois and the win32api. 
- Modify the orgFlag variable on line 9 to work with the company you're trying to filter. This can be found by doing a whois lookup on the domain of one of the spoof phishing emails you've encountered. The variable only needs to include a portion of the organization information.
- The .pyw extension allows for the program to run in the background. The only notification to the user will be if a spoof phishing email is detected. The extension can be changed to .py if you prefer console logging.

### Installing

Software is source and must be run with Python or compiled.

## License

This project is licensed under the GPL License - see the [LICENSE.md](LICENSE.md) file for details
