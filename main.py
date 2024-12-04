import win32com.client
import configparser
import logging
from tqdm import tqdm

# Configure logging
logging.basicConfig(
    filename="email_cleanup.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def get_account(account_name):
    logging.info("Connecting to Outlook...")
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for account in outlook.Folders:
        if account.Name == account_name:
            logging.info(f"Account '{account_name}' found.")
            return account
    logging.error(f"Account '{account_name}' not found.")
    return None

def remove_duplicate_emails(account_name, folder_name="Inbox"):
    logging.info(f"Starting duplicate email removal for account: {account_name}, folder: {folder_name}")
    # Get the specified account
    account = get_account(account_name)
    if not account:
        print(f"Account '{account_name}' not found.")
        return

    # Access the folder
    try:
        folder = account.Folders[folder_name]
        logging.info(f"Folder '{folder_name}' accessed successfully.")
    except Exception as e:
        logging.error(f"Folder '{folder_name}' not found in account '{account_name}'. Error: {e}")
        print(f"Folder '{folder_name}' not found in account '{account_name}'.")
        return

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending
    total_emails = len(messages)

    seen_emails = set()  # Store unique email identifiers
    duplicates = []

    # Iterate through emails
    logging.info("Processing emails...")
    for message in tqdm(messages, desc="Processing emails", total=total_emails):
        try:
            # Create a unique identifier for each email
            identifier = f"{message.Subject}_{message.SenderEmailAddress}_{message.ReceivedTime}"

            if identifier in seen_emails:
                # Mark the message as a duplicate
                duplicates.append(message)
            else:
                # Add the email to the set of seen emails
                seen_emails.add(identifier)
        except Exception as e:
            logging.error(f"Error processing message: {e}")

    # Delete duplicates
    logging.info(f"Found {len(duplicates)} duplicates. Deleting duplicates...")
    for duplicate in tqdm(duplicates, desc="Deleting duplicates"):
        try:
            duplicate.Delete()
            logging.info("Duplicate email deleted.")
        except Exception as e:
            logging.error(f"Error deleting email: {e}")

    logging.info(f"Finished. Removed {len(duplicates)} duplicate emails.")
    print(f"Finished. Removed {len(duplicates)} duplicate emails.")

if __name__ == "__main__":
    # Read config file
    config = configparser.ConfigParser()
    config.read("config.ini")

    # Get parameters from config
    account_name = config["Outlook"].get("account_name", "")
    folder_name = config["Outlook"].get("folder_name", "Inbox")

    # Validate parameters
    if not account_name:
        logging.error("Account name is missing in the configuration file.")
        print("Account name is missing in the configuration file.")
    else:
        remove_duplicate_emails(account_name, folder_name)
