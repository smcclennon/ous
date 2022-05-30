# Office365 User Scraper
Export everyone in your Office 365 organisation into a `.csv` in seconds, straight from your browser, without any admin tools.

This project reverse engineers the Office 365 Outlook webapp API, collecting all users in an Outlook directory via a single API request. All you have to do is obtain a `BaseFolderID` and paste JavaScript code into your browser console. All users within that BaseFolder will be downloaded to a .csv file on your computer.

The `x-owa-canary` cookie is automatically retrieved from your browser and used to authenticate the API request. The API response is then parsed and entered into a 2d array. This array is then converted into a comma-separated-value format which is then downloaded as `user_db.csv` via your browser.

![image](https://user-images.githubusercontent.com/24913281/170899229-8676c592-69e5-4026-9134-61feda1c153f.png)

### Sample csv
Below is the data structure of the `.csv` file generated, formatted as a Markdown table (information redacted):
|Outlook ID| Full Name|Email Address|
|:-:|:-:|:-:|
|AAUQAGAxxxxxxxxxxxxxxxxT2rY=|Aa___ P___|P___@domain.org|
|AAUQAP2xxxxxxxxxxxxxxxxVfSs=|Ab___ E___|E___@domain.org|
|AAUQAPNxxxxxxxxxxxxxxxxoI3w=|Ab___ G___|G___@domain.org|
|AAUQABGxxxxxxxxxxxxxxxxqDhw=|Ac___ C___|C___@domain.org|
|AAUQAI7xxxxxxxxxxxxxxxxHsjQ=|Ad___ R___|a___@domain.org|
|AAUQAHKxxxxxxxxxxxxxxxxqxWY=|AF___ R___|a___@domain.org|
|AAUQAFQxxxxxxxxxxxxxxxxx3MU=|Ai___ N___|N___@domain.org|
|AAUQABdxxxxxxxxxxxxxxxxsi1w=|Al___ D___|D___@domain.org|

## Features
- Retrieve full name, email address and unique Outlook ID by default
- Export all users to a `.csv` file
- Very portable, simply paste code into your browser console
- Quiet network traffic (only 1 request)
- Automatically retrieve the required cookie for API authentication
- Easily extract more information from the API response (see [appendix](#Example-API-response))

## Usage
1. Visit https://outlook.office.com in your browser.
2. Press `F12` to launch the "Developer Tools" popup.
3. Navigate to the "Console" tab within the Developer Tools popup.
4. Paste [this](https://raw.githubusercontent.com/smcclennon/ous/master/scrape-outlook-contacts.js) JavaScript code into the Developer Tools Console (don't execute it yet).
5. Edit the code you just pasted and change `const base_folder_id = ""` so that your BaseFolderId is within the quotation marks. Please see [below](#How-to-get-a-BaseFolderId) on how to get a BaseFolderId.
6. Press `Enter` to execute the code in the Console. Userdata will be printed to the console and downloaded to your computer shortly. If something went wrong, you will receive a JavaScript error in the Console, so make sure your Console is not filtering out errors.

## How to get a BaseFolderId
1. Visit https://outlook.office.com/people/.
2. Press `F12` to launch the "Developer Tools" popup for this tab.
3. In Developer Tools, go to the "Network" tab.
4. Go back to your Outlook browser tab (opened in step 1). On the left you should see a list of user directories (this may be hidden behind the burger menu). Click on the user directory you want to scrape. Many new requests should pop up in your Developer Tools Network tab once you do this.

![image](https://user-images.githubusercontent.com/24913281/170897328-ae7680dd-a036-4d6f-ab38-a45593591fa6.png)

5. On the Developer Tools Network tab, identify the first request that occurred when you completed step 4. The request URL/file should look similar to: `service.svc?action=FindPeople&app=People&n=33`. *(If you find it difficult identifying which request occurred first, try clearing the Network tab request list (bin icon) and then performing step 4 again. The correct request will then most likely be the first one in the list).*

![image](https://user-images.githubusercontent.com/24913281/170897905-2f3b13d0-6e20-4bc8-b185-9fe1d1c84c77.png)

6. Right click that request and select "copy request headers".
7. Paste those request headers into any text editor and then identify the `x-owa-urlpostdata` header.
8. Copy the contents of the `x-owa-urlpostdata` header and paste them into a URL decoder such as: https://www.freeformatter.com/url-encoder.html. *This step isn't necessary, but makes it easier to read the header if you are unfamiliar with URL escape codes.*
9. Copy the decoded header content into any text editor, and identify the `BaseFolderId` key.
10. Copy the `Id` child-key (`["BaseFolderId"]["Id"]`) value. This should look something like `a000a000-0aa0-0a0a-aa00-a000a0000a0a`.
11. The value you just copied is your BaseFolderId.

## Appendix

### Example API response
The API responds with a list of users. Below is the data structure returned per user (some information redacted):
```json
[
    {
      "__type": "PersonaType:#Exchange",
      "PersonaId": {
        "__type": "ItemId:#Exchange",
        "Id": "AAUQAGAxxxxxxxxxxxxxxxxT2rY="
      },
      "PersonaTypeString": "Person",
      "CreationTimeString": "0001-01-02T00:00:00Z",
      "DisplayName": "Aa___ P___",
      "DisplayNameFirstLast": "Aa___ P___",
      "DisplayNameLastFirst": "Aa___ P___",
      "FileAs": "",
      "GivenName": "Aa___",
      "Surname": "P___",
      "CompanyName": "My Domain",
      "EmailAddress": {
        "Name": "Aa___ P___",
        "EmailAddress": "P___@domain.org",
        "RoutingType": "SMTP",
        "MailboxType": "Mailbox"
      },
      "EmailAddresses": [
        {
          "Name": "Aa___ P___",
          "EmailAddress": "P___@domain.org",
          "RoutingType": "SMTP",
          "MailboxType": "Mailbox"
        }
      ],
      "ImAddress": "sip:p___@domain.org",
      "WorkCity": "Watford",
      "RelevanceScore": 2147483647,
      "AttributionsArray": [
        {
          "Id": "0",
          "SourceId": {
            "__type": "ItemId:#Exchange",
            "Id": "AAUQAGAxxxxxxxxxxxxxxxxT2rY="
          },
          "DisplayName": "GAL",
          "IsWritable": false,
          "IsQuickContact": false,
          "IsHidden": false,
          "FolderId": null,
          "FolderName": null,
          "IsGuest": false
        }
      ],
      "ADObjectId": "aa000000-a0aa-00a0-0000-aaa000a0aaa0"
    }
]
```

### x-owa-urlpostdata decoded
`x-owa-urlpostdata` is a header used in the POST request to the Outlook API. We customise the following values in this header:
- `Offset`: Starting index of users to send. An offset of `20` will not return the first 20 users. By default, an offset of `0` is used to return all users.
- `MaxEntriesReturned`: Maximum number of users to be returned by the API. See the [Example API response](#Example-API-response) appendix to view the information returned per user. By default, we request a maximum of `1000` users to be returned. However, you can increase this if you need to.
- `BaseFolderId Id`: This is the Outlook userlist/directory to return in the API response. This is tedious to obtain, but is essential and must be valid or the API request will fail.
```json
{
    "__type": "FindPeopleJsonRequest:#Exchange",
    "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
        "__type": "TimeZoneContext:#Exchange",
        "TimeZoneDefinition": {
            "__type": "TimeZoneDefinitionType:#Exchange",
            "Id": "GMT Standard Time"
        }
        }
    },
    "Body": {
        "IndexedPageItemView": {
        "__type": "IndexedPageView:#Exchange",
        "BasePoint": "Beginning",
        "Offset": Offset,
        "MaxEntriesReturned": MaxEntriesReturned
        },
        "QueryString": null,
        "ParentFolderId": {
        "__type": "TargetFolderId:#Exchange",
        "BaseFolderId": {
            "__type": "AddressListId:#Exchange",
            "Id": BaseFolderId
        }
        },
        "PersonaShape": {
        "__type": "PersonaResponseShape:#Exchange",
        "BaseShape": "Default",
        "AdditionalProperties": [
            {
            "__type": "PropertyUri:#Exchange",
            "FieldURI": "PersonaAttributions"
            },
            {
            "__type": "PropertyUri:#Exchange",
            "FieldURI": "PersonaTitle"
            },
            {
            "__type": "PropertyUri:#Exchange",
            "FieldURI": "PersonaOfficeLocations"
            }
        ]
        },
        "ShouldResolveOneOffEmailAddress": false,
        "SearchPeopleSuggestionIndex": false
    }
}
```