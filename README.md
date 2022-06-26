# Google-Script-Contacts-Sync

This script is intended to synchronize all contacts between multiple Google users' accounts.  It could be modified to share only specific groups, but that is beyond the scope of my own needs.  Please feel free to modify it if you desire to synchronize only specific groups.

I'm not a professional programmer and wouldn't really even call myself a hobbyist.  I needed contacts synchronization and the options at hand were either too expensive, unreliable, or didn't have the features I needed (primarily group synchronization), so I Frankensteined this thing together.  After writing this script, I can see why so many of the synchronization tools available now do not have group synchronization - it was the most difficult part of the script to get to run somewhat reliably.

v2.3 Beta changelog:
- Added a few items to the instructions to provide greater granularity to the setup to hopefully avoid errors.

v2.0 Beta changelog:
- Created queuing sheets for each user.  This allows large changes to be made without causing loop syncing errors.
- Daily Maintenance Function has been created.  If you are starting from scratch, a trigger will be created for you.  If you are upgrading, add a dailyMaintenance trigger every 24 hours.  This will remove deleted contacts from the spreadsheet.
- Category sync'ing issues have been fixed.
- Other minor fixes.
- If you are upgrading from v1.x to v2.x, you need to create an additional sheet for each user in the syncAccounts variable in the Google Sheets workbook.  For example, if syncAccounts is ['email1@gmail.com', 'email2@gmail.com'], create two new sheets labeled _email1@gmail.com_ and _email2@gmail.com_.

Planned changes:
- Review entire code to reduce length and find efficiencies.

See the GSCS.js file for more information and setup instructions.
