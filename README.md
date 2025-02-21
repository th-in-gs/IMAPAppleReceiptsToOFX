# IMAPAppleReceiptsToOFX

Python script to process Apple receipts from an IMAP email account and generate an OFX file for financial software.

## Features

- Connects to an IMAP email account to fetch Apple receipts.
- Searches a specified folder for Apple receipt emails.
- Parses the email content to extract receipt details.
- Generates an OFX file with the receipt transactions.

## Requirements

- Python 3.6+
- `imaplib` for IMAP email access
- `email` for email parsing
- `BeautifulSoup` for HTML parsing
- `moneyed` for currency handling
- `keyring` for secure password storage
- `yaml` for configuration file parsing
- `argparse` for command-line argument parsing


## Configuration

Create a configuration file in YAML format with the following structure:

```yaml
IMAP:
  server: "imap.your-email-provider.com"
  email: "your-email@example.com"
  folder: "Apple Receipts"  # Optional, defaults to "Apple Receipts"
```


## Usage

1. Store your email account password securely using `keyring`:
    ```sh
    python -m keyring set IMAPAppleReceiptsToOFX imap.your-email-provider.com
    ```

2. Run the script with the configuration file, specifying the output file and, optionally, the number of days of receipts to include (default is 90):
    ```sh
    python IMAPAppleReceiptsToOFX.py --config config.yaml --output output.ofx --days 90
    ```

Account and transaction IDs in the OFX are generated from your email address and the order IDs from the receipts, so they should be stable. Your finance software should not create duplicates if you import the same receipts more than once.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.