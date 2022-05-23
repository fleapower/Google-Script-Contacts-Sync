# Google-Script-Contacts-Sync
Synchronize Google contacts across multiple Google accounts.

If you are upgrading from v1.x to v2.x, you need to create an additional sheet for each user in the syncAccounts variable in the Google Sheets workbook.  For example, if syncAccounts is ['email1@gmail.com', 'email2@gmail.com'], create two new sheets labeled _email1@gmail.com_ and _email2@gmail.com_.

v2.3 Beta changelog:
- Added a few items to the instructions to provide greater granularity to setup to hopefully avoid errors.

v2.0 Beta changelog:
- Created queuing sheets for each user.  This allows large changes to be made without causing loop syncing errors.
- Daily Maintenance Function has been created.  If you are starting from scratch, a trigger will be created for you.  If you are upgrading, add a dailyMaintenance trigger every 24 hours.  This will remove deleted contacts from the spreadsheet.
- Category sync'ing issues have been fixed.
- Other minor fixes.

Planned changes:
- Review entire code to reduce length and find efficiencies.

See the GSCS.js file for more information and setup instructions.
