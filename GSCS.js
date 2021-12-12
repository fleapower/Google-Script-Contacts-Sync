/**

GSCS (Google Scripts Contact Sync) Version 1.2 Beta

This script is intended to synchronize all contacts between Google users.  It could easily be modified to share only specific groups, but that is beyond the scope of my own needs.  Please feel free to modify it if you desire to synchronize only specific groups.

I'm not a professional programmer and wouldn't really even call myself a hobbyist.  I needed contacts synchronization and the options at hand were either too expensive, unreliable, or didn't have the features I needed (primarily group synchronization).  After writing this script, I can see why so many of the synchronization tools available now do not have group synchronization - it was the most difficult part of the script to get to run somewhat reliably.


SETUP

Before setting up the script, you need to put all of your contacts into a single, "master" Google account and delete all of the contacts in the "client" accounts.  If you want to keep the contacts in the client accounts, you should export them using Google contacts functions and import them into the master account.  Once your have all of your contacts in a single account, follow these steps:

1)  Go to https://script.google.com while logged in to your master Google account.
2)  Click on the "New Project" button in the top left.
3)  Click on "Untitled Project" in the top left and rename the script.  The name is up to you, but I recommend something like "Google Contacts Sync" so you know what it is later.
4)  Copy and paste this entire script into the editing window (be sure to replace the blank function text in the editing window).
5)  Click the Sharing button in the top right and give Editor access to everyone who will be using the script (you can add more users later - see instructions below.)  Be aware, if you give other users editing privileges, they can edit the script.  If you do not want them to be able to edit the script, you will need to create their own script using steps 1-4.
6)  Change the "syncAccounts" variable below to include the email addresses of the accounts you wish the synchronize.  Note: If you are setting up multiple scripts and not using a shared script, you need to ensure the email addresses are listed in the same order across all scripts.
*/

var syncAccounts = ['email1@gmail.com', 'email2@gmail.com', 'email3@googleAppsDomain.com'];

/**
7)  Go to https://drive.google.com.
8)  Click on "New -> Google Sheets."  This spreadsheet can be located anywhere in any folder you choose and can be named anything you would like (again, I recommend a name easy to remember).
9)  Share the document with the same users from step 5.  Be sure to give them edit access.  Again, nefarious users can destroy the spreadsheet and cause you to lose contacts.
10) Copy the document ID from the URL.  Specifically, you will see something like, "https://docs.google.com/spreadsheets/d/############################################/edit#gid=0."  The document ID will be located where the string of #s are.
11) Paste the document ID of the spreadsheet here:
*/

var ss = SpreadsheetApp.openById('############################################');

/**
12) Go back to your script in your master account.
13) Select "MasterInit" from the function pulldown and click "Run."  Permission will need to be granted for the script to run.  Initialization will take about 1 minute for every 5,000 contacts you have.
14) Next, log into each of the client accounts and view the shared script (or the copied script if you created a separate copy).
15) Select "ClientInit" from the function pulldown and click "Run."  Again, permission will need to be granted.  Because of Google's read/write quotas, this will take a very long time - about an hour for every 1,000 contacts you have in the master acount.  IMPORTANT: DO NOT make changes in any account until the client receives an email that the client initialization is done.  The script is fairly robust in handling errors, but if you make changes (especially deleting contacts), synchronization could be broken for some contacts and the script itself could stop working.  ClientInit should work simultaneously for multiple accounts, but it has not been tested.

All triggers for the script are set by the script itself.


ADD USERS

1)  Add the new email to the end of the list of email addresses in the syncAccounts variable above.
2)  Grant editor access or create a copy of the script itself for the client.
3)  Share the spreadsheet with the client giving edit access.
4)  Run ClientInit from the client's account.


DELETE USERS

To delete users, you can simply delete the trigger running the "syncContacts" function.  However, this will leave deleted contacts in the spreadsheet since this account will now no longer delete those contacts.  A better (and lengthier) way to delete users is to:

1)  Delete all triggers so the scripts don't run while you are doing the following steps:
2)  Open the contacts spreadsheet you created.
3)  Delete the columns of resource IDs and update times at the far right of the data which corresponds to the user being deleted.  If the user is the third listed in the "syncAccounts" variable, you would delete the third pair of corresponding ID and update time columns.
4)  Delete the user account email from "syncAccounts."
5)  Run the "createSyncContactsTrigger" while logged into each account (master and client).

A simpler but even longer way to do the above would be to reinitialize the entire script by following the initial setup instructions (be sure to delete all triggers before you do so).


KNOWN ISSUES
1)  Contact groups do not sync reliably.  The last few versions have improved this immensely, and I haven't seen any problems recently, but I still don't trust it.  I'm still working on this issue.  It is highly recommended you occasionally confirm labels are the same between groups and especially after initial synchronization.
2)  If you need to update more than 300 contacts, do it in increments of less than 300.  For example, if you need to update 400 contacts, do it in two batches of 200 and wait until the first sync cycle is complete (20 minutes) before doing the second 200.  An alternative is to delete all project triggers from all accounts and go through the setup instructions again (you can use the same spreadsheet).
3)  Google is very sensitive about quotas, and the script will occasionally bust a quota.  However, as mentioned earlier, the algorithm is very robust in handling errors.  Synchronization errors should resolve themselves within a iterations depending on the size of the update causing the error.


PROBLEM REPORTING OR TROUBLESHOOTING

If you have problems, please post it on GitHub and I will address it as soon as possible: https://github.com/fleapower/Google-Script-Contacts-Sync/issues.  If possible, include the log from the script's executions page.

*/


var startTime = new Date();

var currUser = Session.getEffectiveUser().getEmail();
Logger.log(currUser)
var currUserNum = syncAccounts.indexOf(currUser);
var otherUserNum = Math.abs(currUserNum - 1)

var sheet = ss.getSheets()[0];
sheet.activate();
var syncTokenFileName = currUser + "_PeopleSyncToken.txt";
var updatedContacts;
var pageSize = 2000;
var masterPersonFields = 'addresses,biographies,birthdays,calendarUrls,clientData,emailAddresses,events,externalIds,genders,imClients,interests,locales,locations,memberships,metadata,miscKeywords,names,nicknames,occupations,organizations,phoneNumbers,relations,sipAddresses,urls,userDefined'

var contactGroupsList = getContactGroupsList();

function MasterInit() {
  RefreshSyncToken();
  deleteAllTriggers();
  var n = 0;
  var appendArray = new Array();
  var updateTime = new Date();
  var allContactsArray = new Array();
  sheet.activate()
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

  // create syncContacts trigger
  createSyncContactsTrigger();
  showElapsedTime();
}

function ClientInit() {
  RefreshSyncToken();
  // since ClientInit is going to create triggers, delete all triggers so they don't overlap/conflict
  deleteAllTriggers();

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
    RefreshSyncToken()
    // and create syncContacts trigger
    createSyncContactsTrigger();
    MailApp.sendEmail({
      to: currUser,
      subject: "GS Contacts Sync is Synchronizing!",
      htmlBody: "The initial contact sync is complete!  You may now modify your contacts as usual."
    })
  }
}

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function createSyncContactsTrigger() {
  ScriptApp.newTrigger("syncContacts")
    .timeBased()
    .everyMinutes(10) // contacts will be synced every hour
    .create();
}

function syncContacts() {
  // Get new contacts from user (otherwise, when contacts are sync'd to spreadsheet, it will start a loop)
  updatedContacts = getUpdatedContacts();
  showElapsedTime();
  SpreadsheetToContacts();
  showElapsedTime();
  ContactsToSpreadsheet(updatedContacts);
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
  var dateArray = new Array();
  var allNewMemberships = new Array();
  var newMembership;
  Logger.log("SpreadsheetToContacts")
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
          Logger.log("     Successfully deleted contact from Google account.")
        }
        catch {
          Logger.log("     Had to use alternate delete.")
        }
        // If doing multiple users, remove contact from spreadsheet only after last user has deleted contact (i.e. all users have same update date/time)
        //    - update dateArray by replacing myUpdate with newestUpdate
        //    - if all dates (use array every function) are now the same in dateArray, then delete contact
        dateArray[currUserNum] = newestUpdate;
        var oldestUpdate = new Date(Math.min.apply(null, dateArray));
        if (newestUpdate.toString() == oldestUpdate.toString()) {
          sheet.deleteRow(i + 1)
          Logger.log ("     Row deleted.")
        }
        else {
          sheet.getRange(i + 1, (currUserNum + 1) * 2 + 25).setValue(newestUpdate)
        }
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
            Logger.log("     Created Group Name: " + groupName)
            Utilities.sleep(2000) // for some reason, Google needs a delay after creating a group before adding a contact or Google can't find the group and will throw an error
          }

          else {
            var groupResourceName = contactGroupsList.map(data => data[1])[groupIndex];
          }

          // ADD GROUP TO MEMBERSHIPS:

          newMembership = {
            "contactGroupMembership" : {
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

        Utilities.sleep(500); // delay is required to not exceed read/write quotas
        
        if ((contactRows[i][(currUserNum + 1) * 2 + 23])) {
          // Update Contact from Spreadsheet
          var people = People.People.getBatchGet({
            resourceNames: [(contactRows[i][(currUserNum + 1) * 2 + 23])],
            personFields: 'metadata'
          });
          if (people.responses[0].person) {
            bodyRequest.etag = people.responses[0].person.etag;
            var personResourceName = (contactRows[i][(currUserNum + 1) * 2 + 23]);
            // Calendar URLs is listed as an update field, but it throws an error when included below.
            Logger.log("     Updating Google contact...")
            People.People.updateContact(bodyRequest, personResourceName, { updatePersonFields: "addresses,biographies,birthdays,clientData,emailAddresses,events,externalIds,genders,imClients,interests,locales,locations,memberships,miscKeywords,names,nicknames,occupations,organizations,phoneNumbers,relations,sipAddresses,urls,userDefined" });
            sheet.getRange(i + 1, (currUserNum + 1) * 2 + 25).setValue(newestUpdate)
          }
          else{
            MailApp.sendEmail({
              to: currUser,
              subject: "GSCS Sync Conflict",
              htmlBody: "Contact " + (contactRows[i][(currUserNum + 1) * 2 + 23]) + " has been deleted, but other users have updated the contact's information since you deleted it.  It is recommend you take this contact out of your trash.  If you still want to delete it, try again after removing it from the trash."
            })
          }
        }
        else {
          // Add Contact from Spreadsheet
          var newContact = People.People.createContact(bodyRequest);
          sheet.getRange(i + 1, (currUserNum + 1) * 2 + 24).setValue(newContact.resourceName)
          sheet.getRange(i + 1, (currUserNum + 1) * 2 + 25).setValue(newestUpdate)
          Logger.log("     Adding Google contact...")
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

function ContactsToSpreadsheet(updatedContacts) {
  Logger.log("ContactsToSpreadsheet")
  var dateArray = new Array();
  var column = (currUserNum + 1) * 2 + 24; //column Index   
  var columnValues2D = sheet.getRange(1, column, sheet.getLastRow()).getValues()
  var columnValues = String(columnValues2D).split(","); // convert to single dimension array
  if (updatedContacts.connections != null) {
    updatedContacts.connections.forEach(function (contact) {
      var updateTime = new Date();
      // find contact in spreadsheet which matches updated contact from Contacts
      searchResult = columnValues.indexOf(contact.resourceName);

      contactArray = getContactArray(contact);

      // remove ID tags and stringify
      contactArray = RemoveIDandStringify(contactArray)

      // Store Contact Groups
      if (contact.memberships != null) {
        var contactGroups = []
        for (var i = 0; i < contact.memberships.length; i++) {
          for (var j = 0; j < contactGroupsList.length; j++) {
            if (contact.memberships[i].contactGroupMembership.contactGroupResourceName == contactGroupsList[j][1]) {
              contactGroups[i] = contactGroupsList[j][0]
            }
          }
        }
        contactArray[13] = contactGroups.join(",") // 14th column is where memberships are; will need to be changed if script updated
      }

      if (searchResult != -1) {
        dateArray = [];
        for (var i = 0; i < syncAccounts.length; i++) {
          dateArray.push(sheet.getRange(searchResult + 1, (i + 1) * 2 + 25).getValue())
        }
        // Update contact

        Logger.log ("     Updating spreadsheet...")
        sheet.getRange(searchResult + 1, 1, 1, 25).setValues([contactArray]);
        sheet.getRange(searchResult + 1, (currUserNum + 1) * 2 + 25).setValue(updateTime);
      }
      else {
        // Add contact

        // append array with individual user's resourceNames
        var resourceNamesArray = [];
        for (var i = 0; i < syncAccounts.length; i++) {
          if (currUser == syncAccounts[i]) {
            resourceNamesArray = resourceNamesArray.concat([contact.resourceName, updateTime]);
          }
          else {
            resourceNamesArray = resourceNamesArray.concat(["", ""]);
          }
        }
        // add contact to end of sheet
        var appendArray = contactArray.concat(resourceNamesArray);
        Logger.log ("     Adding contact to spreadsheet...")
        sheet.appendRow(appendArray);
      }
    })
  }
}

function getContactArray (contact) {
  return ([contact.addresses, contact.biographies, contact.birthdays, contact.calendarUrls, contact.clientData, contact.emailAddresses, contact.events, contact.externalIds, contact.genders, contact.imClients, contact.interests, contact.locales, contact.locations, contact.memberships, contact.metadata, contact.miscKeywords, contact.names, contact.nicknames, contact.occupations, contact.organizations, contact.phoneNumbers, contact.relations, contact.sipAddresses, contact.urls, contact.userDefined]);
}

function SaveContactGroups() {
  if (contact.memberships != null) {
    var contactGroups = []
    for (var i = 0; i < contact.memberships.length; i++) {
      for (var j = 0; j < contactGroupsList.length; j++) {
        if (contact.memberships[i].contactGroupMembership.contactGroupResourceName == contactGroupsList[j][1]) {
          contactGroups[i] = contactGroupsList[j][0]
        }
      }
    }
    contactArray[13] = contactGroups.join(",") // 14th column is where memberships are; will need to be changed if script updated
  }
}

function RefreshSyncToken() {
  Logger.log("RefreshSynctoken")
  var i = 0;
  var pageToken;
  var syncTokenFiles = DriveApp.getFilesByName(syncTokenFileName);
  var syncTokenFile = syncTokenFiles.next();
  var syncToken = syncTokenFile.getBlob().getDataAsString("utf8");

  do {
    var connections = People.People.Connections.list('people/me', {
      pageToken: pageToken,
      requestSyncToken: true,
      syncToken: syncToken,
      pageSize: pageSize,
      personFields: masterPersonFields,
      sources: ["READ_SOURCE_TYPE_CONTACT"]
    });
    try {
      connections.connections.forEach(function (person) {
        // var contactArray = [person.metadata];
        // next page token
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
  Logger.log(newSyncToken)
}

function getUpdatedContacts() {
  Logger.log("getUpdatedContacts")
  var pageToken;

  var syncTokenFiles = DriveApp.getFilesByName(syncTokenFileName);
  var syncTokenFile = syncTokenFiles.next();
  var syncToken = syncTokenFile.getBlob().getDataAsString("utf8");

  do {
    var connections = People.People.Connections.list('people/me', {
      pageToken: pageToken,
      syncToken: syncToken,
      pageSize: pageSize,
      personFields: masterPersonFields,
      sources: ["READ_SOURCE_TYPE_CONTACT"]
    });
  }
  while (pageToken);
  //return(updatedContacts);
  try {
    Logger.log("     Retrieved " + connections.connections.length + " updated contact(s) from user account.")
  }
  catch {
    Logger.log("     Retrieved 0 updated contacts from user account.")
  }
  return (connections);
}

function getContactGroupsList() {
  Logger.log("getContactsGroupsList")
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
