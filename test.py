import win32com.client
import pandas as pd
# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

def filter_job_application_emails(messages): #filters emails 
    """
    Filters a list of email messages to keep only those that contain 
    at least one of the specified keywords in the subject, sender, or body.

    Parameters:
    messages (list): A list of email message objects.
    keywords (list): A list of keywords to search for.

    Returns:
    list: A list of filtered email messages containing at least one keyword.
    """
    filtered_messages = []
    keywords = ["candidatura", "application",  "position"]

    for message in messages:
        
        # Check if any keyword is in the subject, sender name, or body
        if (any(keyword.lower() in message["Subject"].lower() for keyword in keywords) or
            any(keyword.lower() in message["Sender"].lower() for keyword in keywords) or
            any(keyword.lower() in message["Body"].lower() for keyword in keywords)):
            filtered_messages.append(message)

    return filtered_messages


def list_all_accounts():
    try:
        # Access all accounts
        accounts = outlook.Folders
        
        # Print each account name
        print("Accounts in Outlook:")
        for account in accounts:
            print(f"- {account.Name}")
        
        return accounts  # Return the accounts for further processing

    except Exception as e:
        print(f"An error occurred: {e}")

def list_folders(account):
    """List all folders in the specified account."""
    try:
        print(f"Listing folders in account: {account.Name}")
        for folder in account.Folders:
            print(f"  Folder: {folder.Name}")

    except Exception as e:
        print(f"Could not list folders: {e}")

def read_inbox(account_name):
    try:
        # Access all accounts
        accounts = outlook.Folders

        # Loop through the accounts to find the specified one
        for account in accounts:
            if account.Name == account_name:
                print(f"Accessing inbox for account: {account.Name}")

                # List all folders to find the Inbox
                list_folders(account)

                try:
    # Get the Inbox folder for this account
                    inbox = account.Folders["Boîte de réception"]  # Change this if the folder is named differently

                    # Retrieve the last 10 emails
                    messages = inbox.Items
                    message = messages.GetLast()
                    count = 0

                    while message and count < 10000:
                        # Collect email data
                        email_info = {
                            "ReceiverMail": account.Name,
                            "Sender": message.SenderName,
                            "Subject": message.Subject[:100],
                            
                            "Received Time": f"{message.ReceivedTime}",
                            "Body": message.Body[:100] 
                            
                        }
                        email_data.append(email_info)

                        # Print to console (if needed)
                        print(f"Subject: {message.Subject}")
                        print(f"Sender: {message.SenderName}")
                        print(f"Received Time: {message.ReceivedTime}")
                        print(f"Body: {message.Body[:100]}")
                        print("-" * 40)

                        # Get the next message
                        message = messages.GetPrevious()
                        count += 1
                     #filters data to candidatures only :
                    candidatures = filter_job_application_emails(email_data)

                    # Create a DataFrame from the email data
                    df = pd.DataFrame(candidatures)
                   
                    # Write the DataFrame to an Excel file
                    excel_file_path = r"D:\majd\Bureau\emails.xlsx"  # Change the path as needed
                    df.to_excel(excel_file_path, index=False)

                    print(f"Email data written to {excel_file_path}")

                except Exception as e:
                    print(f"Could not access inbox: {e}")
                return  # Exit after finding the account

        print(f"Account '{account_name}' not found.")

    except Exception as e:
        print(f"An error occurred while accessing the inbox: {e}")

# List all accounts
accounts = list_all_accounts()
email_data = []

for account in accounts:
            print(f"Accessing acc : - {account.Name}")
            read_inbox(account.Name)
# Now access the specific account's inbox
account_to_access = "majdsaidani.prodev@gmail.com"  # Replace with the account you want to access
#read_inbox(account_to_access)