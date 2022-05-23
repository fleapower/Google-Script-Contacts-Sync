/**

GSCS (Google Script Contacts Sync) Version 2.3 Beta

This script is intended to synchronize all contacts between Google users.  It could be modified to share only specific groups, but that is beyond the scope of my own needs.  Please feel free to modify it if you desire to synchronize only specific groups.

I'm not a professional programmer and wouldn't really even call myself a hobbyist.  I needed contacts synchronization and the options at hand were either too expensive, unreliable, or didn't have the features I needed (primarily group synchronization), so I Frankensteined this thing together.  After writing this script, I can see why so many of the synchronization tools available now do not have group synchronization - it was the most difficult part of the script to get to run somewhat reliably.


SETUP

Before setting up the script, you need to put all of your contacts into a single, "master" Google account and delete all of the contacts in the "client" accounts.  If you want to keep the contacts in the client accounts, you should export them using Google contacts functions and import them into the master account.  Once your have all of your contacts in a single account, follow these steps:

0)  Yes, step 0, before you do anything else, export your contacts so you have a backup in case something happens to your contacts.  You can also undo changes and restore lost contacts by clicking the gear icon in the top right of the Google contacts page and selecting "Undo changes."  Regarding the latter, it would be a good idea to note the time so you can restore to that point (I name my exported contacts file with the current date and time for ease to help me undo if needed).
1)  Go to https://script.google.com while logged in to your master Google account.
2)  Click on the "New Project" button in the top left.
3)  Click on "Untitled Project" in the top left and rename the script.  The name is up to you, but I recommend something like "Google Contacts Sync" so you know what it is later.
4)  Copy and paste this entire script into the editing window (be sure to replace the blank function text in the editing window).
5)  Click the Sharing button in the top right and give Editor access to everyone who will be using the script (you can add more users later - see instructions below.)  Be aware, if you give other users editing privileges, they can edit the script.  If you do not want them to be able to edit the script, you will need to create their own script using steps 1-4.
6)  Change the "syncAccounts" variable below to include the email addresses of the accounts you wish the synchronize.  Note: If you are setting up multiple scripts and not using a shared script, you need to ensure the email addresses are listed in the same order across all scripts.
*/

var syncAccounts = ['email1@gmail.com', 'email2@yourdomain.com'];

/**
7)  Open a new tab in your browser and go to https://drive.google.com from your master account.
8)  Click on "New -> Google Sheets."  This spreadsheet can be located anywhere in any folder you choose and can be named anything you would like (again, I recommend a name easy to remember).
9)  Share the document with the same users from step 6.  Be sure to give them edit access.  Again, nefarious users can destroy the spreadsheet and cause you to lose contacts.
10) Copy the document ID from the URL.  Specifically, you will see something like, "https://docs.google.com/spreadsheets/d/############################################/edit#gid=0."  The document ID will be located where the string of #s are.
11) Paste the document ID of the spreadsheet here:
*/

var ss = SpreadsheetApp.openById('/############################################');

/**
12) Create an empty text file on your computer and then upload it to your Google drive (be sure the option to automatically convert uploaded files to Google file format is turned OFF).  It does not matter where you put the file in your Google drive.  However, for convenience, just put it in the directory with the spreadsheet:

YOUR_EMAIL_ADDRESS@gmail.com_PeopleSyncToken.txt

Replace the placeholder email address with the email address from your master account.
13) Do step 12 for all of the accounts you plan to sync by using that account's email address.  For client accounts, you can put the file wherever is most convenient.
14) Go back to your script in your master account.
15) Click the plus sign next to "Services" on the left side of the screen.
16) Select "Peopleapi" in the list and click OK.
17) Select "MasterInit" from the function pulldown and click "Run."  Permission will need to be granted for the script to run.  Initialization will take about 1 minute for every 5,000 contacts you have.
18) Next, log into each of the client accounts and view the shared script (or the copied script if you created a separate copy).
19) Select "ClientInit" from the function pulldown and click "Run."  Again, permission will need to be granted.  Because of Google's read/write quotas, this will take a very long time - about an hour for every 1,000 contacts you have in the master acount.  IMPORTANT: DO NOT make changes in any account until the client receives an email that the client initialization is done.  The script is fairly robust in handling errors, but if you make changes (especially deleting contacts), synchronization could be broken for some contacts and the script itself could stop working.  ClientInit should work simultaneously for multiple accounts, but it has not been tested.
20) Set the following variable to the email address where status emails should be sent which the script occasionally generates.
*/

var statusEmail = 'yourstatusemail@gmail.com'

/**

All triggers for the script are set by the script itself.


ADD USERS

1)  Create the synctoken file in the new users Google Drive. (YOUR_EMAIL_ADDRESS@gmail.com_PeopleSyncToken.txt)
2)  Add the new email to the end of the list of email addresses in the syncAccounts variable above.
3)  Grant editor access or create a copy of the script itself for the client.
4)  Share the spreadsheet with the client giving edit access.
5)  Run ClientInit from the client's account.


DELETE USERS

Method 1 (simplest):

1)  Delete the trigger running the "syncContacts" function.
2)  To ensure the user can't make changes to the script or the spreadsheet, you should also remove sharing permissions with that user.  If you no longer have access to that user's account to delete the trigger, the only way to delete a user is to remove these permissions.

Method 2 (more thorough, but more complex):

1)  Delete all triggers for all users, so the scripts don't run while you are doing the following steps:
2)  Open the contacts spreadsheet you created.
3)  Delete the columns of resource IDs and update times at the far right of the data which corresponds to the user being deleted.  If the user is the third listed in the "syncAccounts" variable, you would delete the third pair of corresponding ID and update time columns.
4)  Delete the user account email from "syncAccounts."
5)  Run the "createSyncContactsTrigger" while logged into each account (master and client).

Method 3 (simpler than methods #1 and #2, most thorough, but longer):

1)  Delete all triggers for all users.
2)  Follow the initial setup instructions.


KNOWN ISSUES

1)  Google is very sensitive about quotas, and the script will occasionally bust a quota.  However, as mentioned earlier, the algorithm is very robust in handling errors.  Synchronization errors should resolve themselves after a few executions.
2)  Because Google contacts were never intended to be synchronized between multiple accounts, there is a slight possibility that a contact will not be synchronized if a user updates the contact in the few seconds between their own contacts being updated and a new sync token being generated.  Given typical script runtimes, this is about a 1:300 chance this will happen.  If you notice an unsync'd contact, simply making a small change to the contact (adding a label and then removing it) will force a resync.


PROBLEM REPORTING OR TROUBLESHOOTING

If you have problems, please post it on GitHub and I will address it as soon as possible: https://github.com/fleapower/Google-Script-Contacts-Sync/issues.  If possible, include the log from the script's executions page (please note, your email address may be captured in the log - be sure to use hashtags vice your email address, e.g., #######@gmail.com).

*/

var startTime = new Date();

var currUser = Session.getEffectiveUser().getEmail();
Logger.log(currUser)
var currUserNum = syncAccounts.indexOf(currUser);

var sheet = ss.getSheets()[0];
var mySheet = ss.getSheetByName(currUser);

var syncTokenFileName = currUser + "_PeopleSyncToken.txt";

var pageSize = 2000;
var masterPersonFields = 'addresses,biographies,birthdays,calendarUrls,clientData,emailAddresses,events,externalIds,genders,imClients,interests,locales,locations,memberships,metadata,miscKeywords,names,nicknames,occupations,organizations,phoneNumbers,relations,sipAddresses,urls,userDefined'

var syncTokenExpired = false;

var contactGroupsList = getContactGroupsList();

function MasterInit() {
  try {
    ss.insertSheet(currUser);
  }
  catch {}
  RefreshSyncToken();
  deleteAllTriggers();
  var n = 0;
  var appendArray = new Array();
  var updateTime = new Date();
  var allContactsArray = new Array();
  // Initialize sheet by deleting all rows
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(1, sheet.getLastRow() - 1);
  }
  sheet.getRange("A1:BB1").clear();

  var pageToken;
  do {
    var connections = People.People.Connections.list('people/me', {
      pageToken: pageToken,
      pageSize: pageSize,
      requestSyncToken: true,
      personFields: masterPersonFields,
      sources: ["READ_SOURCE_TYPE_CONTACT"] // we can't update DOMAIN or PROFILE so CONTACT source is specified (must be consisten through all other calls or sync tokens won't work)
    });

    connections.connections.forEach(function (person) {
      var contactArray = [person.addresses, person.biographies, person.birthdays, person.calendarUrls, person.clientData, person.emailAddresses, person.events, person.externalIds, person.genders, person.imClients, person.interests, person.locales, person.locations, person.memberships, person.metadata, person.miscKeywords, person.names, person.nicknames, person.occupations, person.organizations, person.phoneNumbers, person.relations, person.sipAddresses, person.urls, person.userDefined];

      // remove ID tags and stringify
      contactArray = RemoveIDandStringify(contactArray)

      // append array with individual user's resourceNames
      var resourceNamesArray = [];
      for (var i = 0; i < syncAccounts.length; i++) {
        if (currUser == syncAccounts[i]) {
          resourceNamesArray = resourceNamesArray.concat([person.resourceName, updateTime]);
        }
        else {
          resourceNamesArray = resourceNamesArray.concat(["", 0]);
        }
      }

      // add Contact Groups to array to be saved in spreadsheet
      var contactGroups = []
      for (var i = 0; i < person.memberships.length; i++) {
        for (var j = 0; j < contactGroupsList.length; j++) {
          if (person.memberships[i].contactGroupMembership.contactGroupResourceName == contactGroupsList[j][1]) {
            contactGroups[i] = contactGroupsList[j][0]
          }
        }
      }
      contactArray[13] = contactGroups.join(",") // 14th column is where memberships are; will need to be changed if script updated

      // complete apendArray
      appendArray = contactArray.concat(resourceNamesArray);

      // append allContactsArray with new contact
      allContactsArray[n] = appendArray;
      Logger.log(n)
      n++;
      pageToken = connections.nextPageToken;
    });
  }
  while (pageToken);

  // write entire contacts array to spreadsheet
  sheet.getRange(1, 1, allContactsArray.length, appendArray.length).setValues(allContactsArray);

  // Set Synctoken
  var syncTokenFiles = DriveApp.getFilesByName(syncTokenFileName);
  var syncTokenFile = syncTokenFiles.next();
  syncTokenFile.setContent(connections.nextSyncToken);

  createSyncContactsTrigger();

  createDailyMaintenanceTrigger();

  showElapsedTime();
}

function ClientInit() {

  // delete all triggers since ClientInit will create another
  deleteAllTriggers();

  RefreshSyncToken();

  // create ClientInit to run again in seven minutes, one minute after extended runtime error would be thrown
  ScriptApp.newTrigger("ClientInit")
    .timeBased()
    .after(7 * 60 * 1000)
    .create();

  // run SpreadsheetToContacts until initial sync is done

  var completedStatus = SpreadsheetToContacts();

  // if SpreadsheetToContacts fully completes, delete all triggers and create new trigger
  if (completedStatus == "complete") {
    deleteAllTriggers();
    // refreshSyncToken to start syncing normally from this point
    RefreshSyncToken();

    // create user contacts sync queue
    ss.insertSheet(currUser);

    // create file for syncToken storage
    DriveApp.createFile(syncTokenFileName, "");

    // create syncContacts trigger
    createSyncContactsTrigger();

    MailApp.sendEmail({
      to: statusEmail,
      subject: "GS Contacts Sync is Synchronizing!",
      htmlBody: "The initial contact sync for " + currUser + " is complete!  Contacts may be added, modified, or deleted as usual."
    })
  }
}

function deleteAllTriggers() {
  Logger.log ("  deleteAllTriggers")
  var triggers = ScriptApp.getProjectTriggers();
  for (var i in triggers) {
    if (triggers[i].getHandlerFunction() != "dailyMaintenance") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function createSyncContactsTrigger() {
  ScriptApp.newTrigger("syncContacts")
    .timeBased()
    .everyMinutes(10) // contacts will be synced every 10 minutes
    .create();
}

function createDailyMaintenanceTrigger() {
  ScriptApp.newTrigger("dailyMaintenance")
    .timeBased()
    .atHour(3)
    .everyDays(1)  // maintenance will run every 24 hours
    .create();
}

function syncContacts() {
  // Get new contacts from user (otherwise, when contacts are sync'd to spreadsheet, it will start a loop)
  //updatedContacts = GetUpdatedContacts();
  GetUpdatedContacts();
  showElapsedTime();
  // ContactsToSpreadsheet(updatedContacts);
  UpdatesMerge();
  showElapsedTime();
  SpreadsheetToContacts();
  showElapsedTime();
  // it appears refreshing the sync token too quickly after updates does not capture the most recent updates
  // this does leave open the possibility that contacts which are updated during this interval or while the script is running will not be synchronized
  Utilities.sleep(2000);
  RefreshSyncToken();
  showElapsedTime();
}

function RemoveIDandStringify(contactArray) {
  // remove ID tags and stringify
  for (var i = 0; i < contactArray.length; i++) {
    if (contactArray[i]) {
      removeID(contactArray[i])
    }
    if (contactArray[i]) {
      contactArray[i] = JSON.stringify(contactArray[i])
    }
    else {
      contactArray[i] = {}
    }
  }
  return (contactArray)
}

function removeID(obj) {
  for (prop in obj) {
    if (prop === 'id')
      delete obj[prop];
    else if (typeof obj[prop] === 'object')
      removeID(obj[prop]);
  }
}

function SpreadsheetToContacts() {

  Logger.log("SpreadsheetToContacts")

  var dateArray = new Array();
  var allNewMemberships = new Array();
  var newMembership;
  var contactRows = sheet.getDataRange().getValues();

  for (var i = 0; i < contactRows.length; i++) {

    allNewMemberships = [];

    //get latest updateTime
    dateArray = [];
    for (var s = 0; s < syncAccounts.length; s++) {
      dateArray.push(contactRows[i][(s * 2) + 26])
    }
    if (!dateArray[currUserNum]) {
      dateArray[currUserNum] = 0
    }

    var newestUpdate = new Date(Math.max.apply(null, dateArray));
    var myUpdate = contactRows[i][(currUserNum * 2) + 26]

    //compare it to my updateTime
    if (myUpdate == null || newestUpdate > myUpdate) {

      // Delete
      if (contactRows[i][14].includes("\"deleted\":true")) {
        try {
          People.People.deleteContact(contactRows[i][(currUserNum + 1) * 2 + 23]);
          Logger.log("     Successfully deleted contact " + contactRows[i][(currUserNum + 1) * 2 + 23] + " from Google account.")
        }
        catch {
          Logger.log("     Had to use alternate delete.  " + contactRows[i][(currUserNum + 1) * 2 + 23] + " did not exist.")
        }
        sheet.getRange(i + 1, (currUserNum + 1) * 2 + 25).setValue(newestUpdate)
      }

      // Update or Add
      else {
        LabelString = sheet.getRange(i + 1, 14).getValue()
        Labels = LabelString.split(",")

        // Add groups to new/updated contacts groups
        for (var j = 0; j < Labels.length; j++) {
          // CHECK IF GROUP EXISTS:
          var groupName = Labels[j];
          var groupIndex = contactGroupsList.map(data => data[0]).indexOf(groupName);

          // CREATE GROUP IF DOESN'T EXIST:
          if (groupIndex == -1) {
            var groupResource = {
              contactGroup: {
                name: groupName
              }
            }
            group = People.ContactGroups.create(groupResource);
            var groupResourceName = group.resourceName;
            contactGroupsList.push([groupName, groupResourceName])
            Logger.log("     Created group " + groupName + ".")
            Utilities.sleep(2000) // for some reason, Google needs a delay after creating a group before adding a contact or Google can't find the group and will throw an error
          }

          else {
            var groupResourceName = contactGroupsList.map(data => data[1])[groupIndex];
          }

          // ADD GROUP TO MEMBERSHIPS:

          newMembership = {
            "contactGroupMembership": {
              "contactGroupResourceName": groupResourceName
            }
          }

          allNewMemberships = allNewMemberships.concat(newMembership)

        }

        var bodyRequest = {
          "etag": "",
          "addresses": JSON.parse(contactRows[i][0]),
          "biographies": JSON.parse(contactRows[i][1]),
          "birthdays": JSON.parse(contactRows[i][2]),
          "calendarUrls": JSON.parse(contactRows[i][3]),
          "clientData": JSON.parse(contactRows[i][4]),
          "emailAddresses": JSON.parse(contactRows[i][5]),
          "events": JSON.parse(contactRows[i][6]),
          "externalIds": JSON.parse(contactRows[i][7]),
          "genders": JSON.parse(contactRows[i][8]),
          "imClients": JSON.parse(contactRows[i][9]),
          "interests": JSON.parse(contactRows[i][10]),
          "locales": JSON.parse(contactRows[i][11]),
          "locations": JSON.parse(contactRows[i][12]),
          "memberships": allNewMemberships,
          "metadata": JSON.parse(contactRows[i][14]),
          "miscKeywords": JSON.parse(contactRows[i][15]),
          "names": JSON.parse(contactRows[i][16]),
          "nicknames": JSON.parse(contactRows[i][17]),
          "occupations": JSON.parse(contactRows[i][18]),
          "organizations": JSON.parse(contactRows[i][19]),
          "phoneNumbers": JSON.parse(contactRows[i][20]),
          "relations": JSON.parse(contactRows[i][21]),
          "sipAddresses": JSON.parse(contactRows[i][22]),
          "urls": JSON.parse(contactRows[i][23]),
          "userDefined": JSON.parse(contactRows[i][24])
        };

        if ((contactRows[i][(currUserNum + 1) * 2 + 23])) {
          // Update Contact from Spreadsheet
          var people = People.People.getBatchGet({
            resourceNames: [(contactRows[i][(currUserNum + 1) * 2 + 23])],
            personFields: 'metadata'
          });
          if (people.responses[0].person) {
            bodyRequest.etag = people.responses[0].person.etag;
            var personResourceName = (contactRows[i][(currUserNum + 1) * 2 + 23]);
            Utilities.sleep(700); // delay is required to not exceed read/write quotas
            // Calendar URLs is listed as an update field, but it throws an error when included below.
            People.People.updateContact(bodyRequest, personResourceName, { updatePersonFields: "addresses,biographies,birthdays,clientData,emailAddresses,events,externalIds,genders,imClients,interests,locales,locations,memberships,miscKeywords,names,nicknames,occupations,organizations,phoneNumbers,relations,sipAddresses,urls,userDefined" });
            Logger.log("     Updated Google contact " + personResourceName + ".")
            sheet.getRange(i + 1, (currUserNum + 1) * 2 + 25).setValue(newestUpdate)
          }
          else {
            MailApp.sendEmail({
              to: statusEmail,
              subject: "GSCS Sync Conflict",
              htmlBody: "Contact " + (contactRows[i][(currUserNum + 1) * 2 + 23]) + " for " + currUser + " has been deleted, but other users have updated the contact's information since it was deleted.  It is recommended you remove this contact from the trash.  If this contact should be deleted, try again after removing it from the trash."
            })
          }
        }
        else {
          // Add Contact from Spreadsheet
          var newContact = People.People.createContact(bodyRequest);
          sheet.getRange(i + 1, (currUserNum + 1) * 2 + 24).setValue(newContact.resourceName)
          sheet.getRange(i + 1, (currUserNum + 1) * 2 + 25).setValue(newestUpdate)
          Logger.log("     Adding Google contact " + newContact.resourceName + ".")
        }
      }
    }
    // if function has been running longer than five minutes, return incomplete for ClientInit;  this will keep the script from throwing a max runtime error at six minutes
    if (((new (Date) - startTime) / 1000) > 300) {
      return (i);
    }
  }//);
  return ("complete");
}

function UpdatesMerge() {
  Logger.log("UpdatesMerge")
  var dateArray = new Array();
  var column = (currUserNum + 1) * 2 + 24; //column Index   
  var columnValues2D = sheet.getRange(1, column, sheet.getLastRow()).getValues()
  var columnValues = String(columnValues2D).split(","); // convert to single dimension array

  if (mySheet.getLastRow() > 0) {
    var myUpdates = new Array();
    myUpdates = mySheet.getDataRange().getValues();
    var c = 0;
    do {
      var updateTime = new Date();
      // find contact in spreadsheet which matches updated contact from Contacts
      searchResult = columnValues.indexOf(myUpdates[c][25]);
      var resourceName = myUpdates[c].pop();
      contactArray = myUpdates[c];

      if (searchResult != -1) {
        dateArray = [];
        for (var i = 0; i < syncAccounts.length; i++) {
          dateArray.push(sheet.getRange(searchResult + 1, (i + 1) * 2 + 25).getValue())
        }
        // Update contact

        Logger.log("     Updated " + resourceName + " on spreadsheet.")
        sheet.getRange(searchResult + 1, 1, 1, 25).setValues([contactArray]);
        sheet.getRange(searchResult + 1, (currUserNum + 1) * 2 + 25).setValue(updateTime);
      }
      else {
        // Add contact

        // append array with individual user's resourceNames
        var resourceNamesArray = [];
        for (var i = 0; i < syncAccounts.length; i++) {
          if (currUser == syncAccounts[i]) {
            resourceNamesArray = resourceNamesArray.concat([resourceName, updateTime]);
          }
          else {
            resourceNamesArray = resourceNamesArray.concat(["", ""]);
          }
        }
        // add contact to end of sheet
        var appendArray = contactArray.concat(resourceNamesArray);
        Logger.log("     Added " + resourceName + " to spreadsheet.")
        sheet.appendRow(appendArray);
      }
      c++
    }
    while ((c < myUpdates.length) && ((new (Date) - startTime) / 1000 < 300))
  }

  if (c == mySheet.getLastRow()) {
    mySheet.clear();
  }
  else if (c) {
    mySheet.deleteRows(1, c);
  }

}

function getContactArray(contact) {
  return ([contact.addresses, contact.biographies, contact.birthdays, contact.calendarUrls, contact.clientData, contact.emailAddresses, contact.events, contact.externalIds, contact.genders, contact.imClients, contact.interests, contact.locales, contact.locations, contact.memberships, contact.metadata, contact.miscKeywords, contact.names, contact.nicknames, contact.occupations, contact.organizations, contact.phoneNumbers, contact.relations, contact.sipAddresses, contact.urls, contact.userDefined]);
}

function RefreshSyncToken() {
  Logger.log("RefreshSynctoken")
  var i = 0;
  var pageToken;
  Logger.log(syncTokenFileName);
  var syncTokenFiles = DriveApp.getFilesByName(syncTokenFileName);
  var syncTokenFile = syncTokenFiles.next();
  var syncToken = syncTokenFile.getBlob().getDataAsString("utf8");

  do {
    if (!syncTokenExpired) {
      var connections = People.People.Connections.list('people/me', {
        pageToken: pageToken,
        requestSyncToken: true,
        syncToken: syncToken,
        pageSize: pageSize,
        personFields: masterPersonFields,
        sources: ["READ_SOURCE_TYPE_CONTACT"]
      });
    }
    else {
      var connections = People.People.Connections.list('people/me', {
        pageToken: pageToken,
        requestSyncToken: true,
        pageSize: pageSize,
        personFields: masterPersonFields,
        sources: ["READ_SOURCE_TYPE_CONTACT"]
      });
    }
    try {
      connections.connections.forEach(function (person) {
        var newSyncToken = connections.nextSyncToken;
      });
    }
    catch { }
    pageToken = connections.nextPageToken;
  }
  while (pageToken);

  // Set Synctoken

  var newSyncToken = connections.nextSyncToken;
  syncTokenFile.setContent(newSyncToken);
  return (newSyncToken);
}

function GetUpdatedContacts() {
  Logger.log("GetUpdatedContacts")
  var pageToken;
  var updatedContactsArray = new Array();
  var appendArray = new Array();
  var n = 0;
  var syncTokenFiles = DriveApp.getFilesByName(syncTokenFileName);
  var syncTokenFile = syncTokenFiles.next();
  var syncToken = syncTokenFile.getBlob().getDataAsString("utf8");
  var refreshedSyncToken = new Boolean();

  do {
    try {
      refreshedSyncToken = false;
      var connections = People.People.Connections.list('people/me', {
        pageToken: pageToken,
        syncToken: syncToken,
        pageSize: pageSize,
        personFields: masterPersonFields,
        sources: ["READ_SOURCE_TYPE_CONTACT"]
      });
    }
    catch {
      Logger.log("Sync token error.  Refreshing token.");
      syncTokenExpired = true;
      syncToken = RefreshSyncToken();
      syncTokenExpired = false;
      refreshedSyncToken = true;
      MailApp.sendEmail({
        to: statusEmail,
        subject: "GS Contacts Sync syncToken error!",
        htmlBody: "There was a syncToken error while synchronizing.  Synchronizations may have been lost.  Please check your contacts."
      })
    }

    try {
      connections.connections.forEach(function (person) {
        var contactArray = [person.addresses, person.biographies, person.birthdays, person.calendarUrls, person.clientData, person.emailAddresses, person.events, person.externalIds, person.genders, person.imClients, person.interests, person.locales, person.locations, person.memberships, person.metadata, person.miscKeywords, person.names, person.nicknames, person.occupations, person.organizations, person.phoneNumbers, person.relations, person.sipAddresses, person.urls, person.userDefined];

        contactArray = RemoveIDandStringify(contactArray)

        // add Contact Groups to array to be saved in spreadsheet
        var contactGroups = []
        try {
          for (var i = 0; i < person.memberships.length; i++) {
            for (var j = 0; j < contactGroupsList.length; j++) {
              if (person.memberships[i].contactGroupMembership.contactGroupResourceName == contactGroupsList[j][1]) {
                contactGroups[i] = contactGroupsList[j][0]
              }
            }
          }
        }
        catch { }

        contactArray[13] = contactGroups.join(",") // 14th column is where memberships are; will need to be changed if script updated

        appendArray = contactArray.concat(person.resourceName);

        updatedContactsArray[n] = appendArray;
        n++;
      })
      pageToken = connections.nextPageToken;
    }
    catch { }


  }
  while (pageToken || refreshedSyncToken == true);

  if (updatedContactsArray.length > 0) {
    Logger.log("     " + updatedContactsArray.length + " contacts queued for update.")
    mySheet.getRange(mySheet.getLastRow() + 1, 1, updatedContactsArray.length, appendArray.length).setValues(updatedContactsArray);
  }
}

function getContactGroupsList() {
  Logger.log("getContactGroupsList")
  var contactGroupsListJSON = People.ContactGroups.list({
    "groupFields": "name",
    "pageSize": 1000 // no more than 1000 contact groups
  })
  var contactGroupsList = []
  contactGroupsListJSON.contactGroups.forEach(function (contactGroups) {
    contactGroupsList.push([contactGroups.name, contactGroups.resourceName])
  });
  return (contactGroupsList)
}

function showElapsedTime() {
  executionTime = (new (Date) - startTime) / 1000
  Logger.log("     Elapsed Time: " + executionTime + " seconds.")
}

function dailyMaintenance() {
  Logger.log("dailyMaintenance")
  var rowsDeleted = 0;
  // remove all users' permissions on spreadsheets
  for (var i = 0; i < syncAccounts.length; i++) {
    if (syncAccounts[i] != currUser) {
      ss.removeEditor(syncAccounts[i])
    }
  }

  // keep user triggers from running
  deleteAllTriggers();

  // delete any deleted contact rows from spreadsheet with "deleted":"true
  var column = 15; // 15th column is where '"delete": true' is stored
  var deleteTrueColumns = sheet.getRange(1, column, sheet.getLastRow()).getValues()
  var dateArray = new Array();
  var contactRows = sheet.getDataRange().getValues();

  for (var i = deleteTrueColumns.length - 1; i > 0; i--) {
    for (var j = 0; j < deleteTrueColumns[i].length; j++) {
      if (deleteTrueColumns[i][j].includes("\"deleted\":true")) {
        dateArray = [];
        for (var s = 0; s < syncAccounts.length; s++) {
          if (contactRows[i][(s * 2) + 26].length > 0) { // if update time is missing, then don't add it
            dateArray.push(contactRows[i][(s * 2) + 26])
          }
        }
        var oldestUpdate = new Date(Math.min.apply(null, dateArray));
        var newestUpdate = new Date(Math.max.apply(null, dateArray));
        // only delete if all updte times are the same or missing (they've either all been deleted or they've never been added)
        if (newestUpdate.toString() == oldestUpdate.toString()) {
          sheet.deleteRow(i + 1);
          rowsDeleted++;
        }
      }
    }
  }
  Logger.log("     Rows deleted: " + rowsDeleted)

  // recreate triggers for master account
  createSyncContactsTrigger();
  // createDailyMaintenanceTrigger();

  // restore users' permissions for spreadsheet

  for (var i = 0; i < syncAccounts.length; i++) {
    if (syncAccounts[i] != currUser) {
      ss.addEditor(syncAccounts[i])
    }
  }
}
