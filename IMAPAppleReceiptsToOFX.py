import argparse
import email
import imaplib
import logging
import re
import sys
from collections import Counter, OrderedDict
from datetime import datetime, timedelta
from email.header import decode_header

import keyring
import yaml
from bs4 import BeautifulSoup
from moneyed import Money, USD

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(lineno)d: %(asctime)s - %(levelname)s - %(message)s')

def login_to_imap(imap_server, email_account, password):
    try:
        mail = imaplib.IMAP4_SSL(imap_server)
        mail.login(email_account, password)
        logging.info('Logged in to IMAP server')
        return mail
    except Exception as e:
        logging.error(f'Failed to login to IMAP server: {e}')
        return None

def list_folders(mail):
    try:
        status, folders = mail.list()
        if status == 'OK':
            logging.info('Available folders:')
            for folder in folders:
                logging.info(folder)
        else:
            logging.error('Failed to list folders')
    except Exception as e:
        logging.error(f'Failed to list folders: {e}')

def fetch_emails(mail, folder, days):
    try:
        logging.info(f'Selecting folder: "{folder}"')
        mail.select(f'"{folder}"')
        logging.info(f'Folder Selected: "{folder}"')
        date = (datetime.now() - timedelta(days=days)).strftime("%d-%b-%Y")
        status, messages = mail.search(None, f'(SINCE {date} SUBJECT "Your receipt from Apple.")')
        email_ids = messages[0].split()
        logging.info(f'Fetched {len(email_ids)} emails with subject "Your receipt from Apple." from the last {days} days in folder {folder}')
        return email_ids
    except Exception as e:
        logging.error(f'Failed to fetch emails: {e}')
        return []

def process_email(mail, email_id):
    try:
        status, msg_data = mail.fetch(email_id, '(RFC822)')
        if status != 'OK':
            logging.error(f'Failed to fetch email {email_id}')
            return None

        msg = email.message_from_bytes(msg_data[0][1])
        subject, encoding = decode_header(msg['Subject'])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding if encoding else 'utf-8')
        logging.info(f'Processing email with subject: {subject}')

        receipt_date = msg['Date']
        if receipt_date:
            try:
                receipt_date = datetime.strptime(receipt_date, '%a, %d %b %Y %H:%M:%S %z')
            except ValueError:
                receipt_date = datetime.strptime(receipt_date, '%a, %d %b %Y %H:%M:%S %z (%Z)')
        else:
            logging.error(f'Failed to extract sent date from email {email_id}')
            return None

        recipient_email = msg['To']
        if recipient_email:
            recipient_email = email.utils.parseaddr(recipient_email)[1]
        else:
            logging.error(f'Failed to extract recipient email from email {email_id}')
            return None

        html_content = None
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_disposition = str(part.get("Content-Disposition"))
                if "attachment" not in content_disposition and content_type == "text/html":
                    html_content = part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
                    break
        else:
            if msg.get_content_type() == "text/html":
                html_content = msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8')

        if html_content:
            apple_id = ''
            receipt_items = OrderedDict()
            receipt_subtotal = Money(0, USD)
            receipt_tax = Money(0, USD)
            receipt_total = Money(0, USD)
            receipt_id = ''

            logging.info("HTML content extracted from email")
            soup = BeautifulSoup(html_content, 'html.parser')

            desktop_div = soup.find('div', class_='aapl-desktop-div')
            if desktop_div:
                try:
                    account_label_node = desktop_div.find(string=lambda text: re.search(r"APPLE\s(ACCOUNT|ID)", text))
                    if account_label_node:
                        account_node = account_label_node.find_parent('td')
                        potential_apple_id = account_node.find_all(string=True)[-1].get_text(strip=True).split()[-1]
                        if '@' in potential_apple_id:
                            if any(c.isspace() for c in potential_apple_id):
                                logging.error(f'Parsed Apple ID contains whitespace: {potential_apple_id}')
                            else:
                                apple_id = potential_apple_id.strip()
                        else:
                            logging.error(f'Parsed Apple ID does not contain "@": {potential_apple_id}')
                except (ValueError, AttributeError) as e:
                    logging.error(f'Error parsing Apple ID {e}')

                try:
                    order_label_node = desktop_div.find(string=lambda text: re.search(r"ORDER\sID", text))
                    if order_label_node:
                        order_node = order_label_node.find_parent('td')
                        potential_order_id = order_node.find_all(string=True)[-1].get_text(strip=True).split()[-1]
                        if potential_order_id:
                            if any(c.isspace() for c in potential_order_id):
                                logging.error(f'Parsed Order ID contains whitespace: {potential_order_id}')
                            else:
                                order_id = potential_order_id.strip()
                        else:
                            logging.error(f'Parsed Order ID is empty: {potential_order_id}')
                except (ValueError, AttributeError) as e:
                    logging.error(f'Error parsing Order ID: {e}')

                logging.info(f'Apple ID: {apple_id}')
                logging.info(f'Order ID: {order_id}')

                for item_link in desktop_div.find_all('a', class_='item-links'):
                    cell = item_link.find_parent('td')
                    if cell:
                        item_details = {}

                        for class_name in ['title', 'renewal', 'duration']:
                            span = cell.find('span', class_=class_name)
                            if span:
                                text = span.get_text(strip=True)
                                if text:
                                    item_details[class_name] = text

                        title = item_details['title']
                        if title:

                            # Ick. Why does Apple not name this more explicitly?
                            if title == "Premier (Automatic Renewal)":
                                title = "Apple One Premier"


                            row = cell.find_parent('tr')
                            if(row):
                                price = Money(row.find_all('td')[-1].get_text(strip=True).replace('$', ''), USD)

                                if title.startswith("Money added to"):
                                    price = -price

                                item_details['price'] = price

                            receipt_items[title] = item_details

                def extract_amount_from_div(div, label):
                    try:
                        cell = div.find('td', string=lambda text: text and label.lower() == text.strip().lower())
                        if cell:
                            row = cell.find_parent('tr')
                            return Money(row.find_all('td')[-1].get_text(strip=True).replace('$', ''), USD)
                    except (ValueError, AttributeError) as e:
                        logging.error(f'Error parsing {label}: {e}')
                    return Money(0, USD)

                # Extract subtotal, receipt_total, and receipt_tax amounts
                receipt_subtotal = extract_amount_from_div(desktop_div, 'Subtotal')
                receipt_total = extract_amount_from_div(desktop_div, 'Total')
                receipt_tax = extract_amount_from_div(desktop_div, 'Tax')

                if receipt_tax == Money(0, USD) and receipt_subtotal == Money(0, USD):
                    receipt_subtotal = receipt_total

                # Validate receipt items and totals
                calculated_subtotal = sum(item['price'] for item in receipt_items.values())
                calculated_total = calculated_subtotal + receipt_tax

                if calculated_subtotal != receipt_subtotal:
                    logging.error(f'Subtotal mismatch: calculated {calculated_subtotal}, expected {receipt_subtotal}')
                if calculated_total != receipt_total:
                    logging.error(f'Total mismatch: calculated {calculated_total}, expected {receipt_total}')
            else:
                logging.info("No HTML content found in email")
        else:
            logging.error("Failed to extract HTML content from email")

        if receipt_items:
            item_names = ', '.join(receipt_items.keys())
            if len(receipt_items) > 1:
                logging.info(f'Multiple items found: {item_names}')
            else:
                logging.info(f'Single item found: {item_names}')
        else:
            logging.info('No items found in receipt')

        if (receipt_items and receipt_total):
            return {
                'order_id': order_id,
                'apple_id': apple_id,
                'receipt_items': receipt_items,
                'subtotal': receipt_subtotal,
                'receipt_tax': receipt_tax,
                'receipt_total': receipt_total,
                'date': receipt_date,
                'recipient_email': recipient_email
            }
        else:
            return None

    except Exception as e:
        logging.error(f'Failed to process email {email_id}: {e}')
        return None

def generate_ofx_output(receipt_data, account_id, output_file):
    logging.info('Generating OFX output')

    ofx_data = """
<OFX>
  <SIGNONMSGSRSV1>
    <SONRS>
      <STATUS>
        <CODE>0</CODE>
        <SEVERITY>INFO</SEVERITY>
      </STATUS>
      <DTSERVER>{datetime}</DTSERVER>
      <LANGUAGE>ENG</LANGUAGE>
    </SONRS>
  </SIGNONMSGSRSV1>
  <BANKMSGSRSV1>
    <STMTTRNRS>
      <TRNUID>1001</TRNUID>
      <STATUS>
        <CODE>0</CODE>
        <SEVERITY>INFO</SEVERITY>
      </STATUS>
      <STMTRS>
        <CURDEF>USD</CURDEF>
        <BANKACCTFROM>
          <BANKID>{bank_id}</BANKID>
          <ACCTID>{account_id}</ACCTID>
          <ACCTTYPE>CHECKING</ACCTTYPE>
        </BANKACCTFROM>
        <BANKTRANLIST>
          {transactions}
        </BANKTRANLIST>
        <LEDGERBAL>
          <BALAMT>0.00</BALAMT>
          <DTASOF>{datetime}</DTASOF>
        </LEDGERBAL>
      </STMTRS>
    </STMTTRNRS>
  </BANKMSGSRSV1>
</OFX>
    """

    transaction_template = """
<STMTTRN>
  <TRNTYPE>{type}</TRNTYPE>
  <DTPOSTED>{date}</DTPOSTED>
  <TRNAMT>{amount}</TRNAMT>
  <FITID>{fitid}</FITID>
  <NAME>{title}</NAME>
  <MEMO>{memo}</MEMO>
</STMTTRN>
    """

    transactions = []
    total_receipts = 0
    total_items = 0

    for receipt in receipt_data:
        receipt_items = receipt['receipt_items']
        receipt_subtotal = receipt['subtotal']
        receipt_tax = receipt['receipt_tax']
        receipt_total = receipt['receipt_total']
        receipt_date = receipt['date']
        order_id = receipt['order_id']
        order_apple_id = receipt['apple_id']

        if not receipt_items:
            logging.warning(f'No receipt items found for order ID {order_id}')
            continue

        tax_percentage = receipt_tax / receipt_total

        item_amounts = []
        for title, item in receipt_items.items():
            item_price = item['price']
            item_tax = (item_price * tax_percentage).round(2)
            amount = item_price + item_tax
            item_amounts.append((title, item, amount))

        # Adjust for rounding differences
        total_calculated = sum(amount for _, _, amount in item_amounts)
        rounding_difference = (receipt_total - total_calculated).round(2)
        if item_amounts:
            last_item = item_amounts[-1]
            item_amounts[-1] = (last_item[0], last_item[1], last_item[2] + rounding_difference)

        # Verify that all item amounts add up to the total
        total_calculated = sum(amount for _, _, amount in item_amounts)
        if total_calculated != receipt_total:
            logging.error(f'Total mismatch for receipt on {receipt_date}: calculated {total_calculated}, expected {receipt_total}')
            continue

        memo = ''
        if order_apple_id != account_id:
            memo += f'Apple ID: {order_apple_id}'
        if item.get('duration'):
            if memo:
                memo += '; '
            memo += f'Subscription: {item["duration"]}'

        item_counter = 1
        for title, item, amount in item_amounts:
            transaction = transaction_template.format(
                type='CREDIT' if amount < Money(0, USD) else 'DEBIT',
                date=receipt_date.strftime("%Y%m%d"),
                amount=f"{-amount}",
                fitid=f"{order_id}-{item_counter}",
                title=title,
                memo=memo
            )
            transactions.append(transaction)
            item_counter += 1

        total_receipts += 1
        total_items += len(item_amounts)
        logging.info(f'Processed {len(item_amounts)} items for order ID {order_id}')

    most_recent_date = max(receipt['date'] for receipt in receipt_data).strftime("%Y%m%d%H%M%S")
    ofx_output = ofx_data.format(
        datetime=most_recent_date,
        bank_id="IMAPAppleReceiptsToOFX",
        account_id=account_id,
        transactions="\n".join(transactions)
    )

    with open(output_file, "w") as ofx_file:
        ofx_file.write(ofx_output)

    logging.info(f'OFX output generated with {total_receipts} receipts and {total_items} items')

def main():
    parser = argparse.ArgumentParser(description='Process IMAP Apple Receipts to OFX.')
    parser.add_argument('--config', required=True, help='Path to the config file')
    parser.add_argument('--output', required=True, help='Path to the output OFX file')
    parser.add_argument('--days', type=int, default=90, help='Number of days of receipts to include')
    args = parser.parse_args()

    # Load configuration from the specified file
    with open(args.config, 'r') as config_file:
        config = yaml.safe_load(config_file)

    print(config)

    imap_server = config['IMAP']['server']
    email_account = config['IMAP']['email']
    folder = config['IMAP'].get('folder', 'Apple Receipts')

    logging.info(f'IMAP Server: {imap_server}')
    logging.info(f'Email Account: {email_account}')
    logging.info(f'Folder: {folder}')

    password = keyring.get_password("IMAPAppleReceiptsToOFX", imap_server)

    if password is None:
        logging.error(f'No password found in keychain for {email_account}')
        exit(1)

    mail = login_to_imap(imap_server, email_account, password)
    if mail:
        all_receipt_data = []
        email_ids = fetch_emails(mail, folder=folder, days=args.days)
        recipient_emails = []

        for email_id in email_ids:
            receipt_data = process_email(mail, email_id)
            if receipt_data:
                all_receipt_data.append(receipt_data)
                recipient_emails.append(receipt_data['recipient_email'])

        mail.logout()

        logging.info(f'Total number of receipts: {len(all_receipt_data)}')
        total_items = sum(len(receipt['receipt_items']) for receipt in all_receipt_data)
        logging.info(f'Total number of items: {total_items}')

        if recipient_emails:
            most_common_email = Counter(recipient_emails).most_common(1)[0][0]
            logging.info(f'Most common recipient email: {most_common_email}')
            generate_ofx_output(all_receipt_data, most_common_email, args.output)
        else:
            logging.error('No recipient emails found. Exiting.')

if __name__ == '__main__':
    main()