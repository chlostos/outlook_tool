import win32com.client
import configparser

def get_account(account_name):
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    for account in outlook.Folders:
        if account.Name == account_name:
            return account
    return None

def remove_duplicate_emails(account_name, folder_name="Inbox"):
    # Get the specified account
    account = get_account(account_name)
    if not account:
        print(f"Account '{account_name}' not found.")
        return

    # Access the folder
    try:
        folder = account.Folders[folder_name]
    except Exception as e:
        print(f"Folder '{folder_name}' not found in account '{account_name}'.")
        return

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort by received time, descending

    seen_emails = set()  # Store unique email identifiers
    duplicates = []

    # Iterate through emails
    for message in messages:
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
            print(f"Error processing message: {e}")

    # Delete duplicates
    for duplicate in duplicates:
        try:
            duplicate.Delete()
            print("Duplicate email deleted.")
        except Exception as e:
            print(f"Error deleting email: {e}")

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
        print("Account name is missing in the configuration file.")
    else:
        remove_duplicate_emails(account_name, folder_name)
