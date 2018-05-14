# https://stackoverflow.com/questions/26322255/parsing-outlook-msg-files-with-python?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
import win32com.client
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
msg = outlook.OpenSharedItem(r"./examples/SI-81772.msg")

print(msg.SenderName)
print(msg.SenderEmailAddress)
print(msg.SentOn)
print(msg.To)
print(msg.CC)
print(msg.BCC)
print(msg.Subject)
print(msg.Body)

count_attachments = msg.Attachments.Count
if count_attachments > 0:
    for item in range(count_attachments):
        print(msg.Attachments.Item(item + 1).Filename)

del outlook, msg