# Missing Recipients Reminder Outlook 2016 add-in

In case you forget to add a recipient to an e-mail message, reminds you and gives you a chance to go back and add the missing recipients.

This is done fully on the client (no data is sent on the network) by looking for any of the patterns below in the beginning of your message:

```
+{name}, {...}
Adding {name}, {...}
Added {name}, {...}
```

Such content can also be enclosed within parentheses / square brackets. If any names don't appear in your list of recipients (To, cc, bcc, either in the e-mail address or contact name) you will get the reminder.

Source code available at https://github.com/davidni/MissingRecipientsReminder