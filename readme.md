# Manage-OutlookInbox
## Summary
Provides a command line interface to search and manage emails in your inbox.

The idea for this module came about when my on-prem inbox filled up (2GB limit) and I had to purge a metric ****-ton of emails quickly. I quickly realized Outlook and OWA search were not going to work for me here to do what I needed to do as quickly as I needed to do it. This tool allows for a way to organize and delete email based on very specific search criteria. The only way it could possibly be better is if you could use SQL like syntax with it - which you can't - so don't ask.

### Connecting to O365 using OAuth
(Only supported method. There is an on prem version of this script that uses BASIC auth)

```
Manage-OutlookInbox -newtoken
```

This opens an ADAL window to sign you into Office 365 and get an auth token. This token is stored in the module directory and is only valid for a limited time (about 8 hours). After the token expires, re-run this command to sign in again.

### Searching items
Examples:

`Manage-OutlookInbox -searchQuery "from:someone@example.com"`

`Manage-OutlookInbox -searchQuery "subject:something"`

`Manage-OutlookInbox -searchQuery "broad search term"`

```
Manage-OutlookInbox -searchQuery "subject:`"something here`"" from:specific@person.com"`
```

Example searching a folder other than "Inbox":
```
Manage-OutlookInbox -searchQuery "subject:`"something here`"" from:specific@person.com"` -RootSearchFolderName "Projects"
```

This uses `Out-Gridview` to preview your search results before using an action on them like `-HardDelete`

Note: The search does __NOT__ recurse by design. This is so you can organize your emails into subfolders without affecting other subfolders you already organized inadvertently.

### Doing things with results
Specify `-HardDelete` after to delete everything in search permanently

Specify `-SoftDelete` after to delete everything in search permanently but still be recoverable

Specify `-DeleteToDeletedItems` after to move to "Deleted Items" in Outlook

Example:
```
Manage-OutlookInbox -searchQuery "subject:`"something here`"" from:specific@person.com"` -HardDelete
```

### Move results to a folder (FOLDER NAME MUST BE UNIQUE)
In the interest of ease-of-use, this command searches for the destination folder by name. If more than one match is found it errors out. If need-be, temporarily rename your folder in Outlook to use this command and then name it back when you're done.

Example:
```
Manage-OutlookInbox -searchQuery "subject:`"something here`"" from:specific@person.com"` -MoveToFolder "Some Folder Name"
```

### BONUS!
Clean up pesky calendar invite emails without deleting them from your calendar! These stay in your inbox and are annoying to look at because they are on your calendar! Let's get rid of them!

To view:
```
Manage-OutlookInbox -CleanupCalendarInvitesInInbox -List
```

To delete:
```
Manage-OutlookInbox -CleanupCalendarInvitesInInbox -SoftDelete
Manage-OutlookInbox -CleanupCalendarInvitesInInbox -HardDelete
Manage-OutlookInbox -CleanupCalendarInvitesInInbox -DeleteToDeletedItems
```


#### Note
There are probably a few parameters left over in here from before I converted it from the on-prem verion to the OAuth version. Ignore these. I haven't tested the "Mailbox" cmdlet yet on 365.