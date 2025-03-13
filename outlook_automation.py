import os
import win32com.client

# Path to the folder where attachment will be saved
folder_path = r"C:\Users\xxx"
save_filename = "file" # Fixed filename for the attachment

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Select the Inbox folder
mailbox = outlook.Folders("xyz@abc.com")
folder_inbox = mailbox.Folders("Inbox")

# Get all emails sorted by received time (newest first)
all_email = folder_inbox.Items
all_email.Sort("[ReceivedTime]", True)

# Check if the email is from the correct sender and has the expected subject
for message in all_email:
    if message.SenderEmailAddress == "sender@email.com" and message.Subject.startswith("subject"):
        # Download attachment
        for attachment in message.Attachments:
            file_extension = os.path.splitext(attachment.FileName)[1]  # Keep original file extension
            save_path = os.path.join(folder_path, save_filename + file_extension)
            attachment.SaveAsFile(save_path)
            print(f"Saved: {save_path}")
        break  # Stop after saving the first attachment