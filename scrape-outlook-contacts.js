// Scrape all users on your Office 365 Outlook people directory
// github.com/smcclennon

// Delete all variables created in this block once execution finishes
(async () => {

    // These need to be replaced with valid values or the API request will fail
    const base_folder_id = "a000a000-0aa0-0a0a-aa00-a000a0000a0a"

    // Download text to a file
    // https://stackoverflow.com/a/47359002
    function saveAs(text, filename){
        var pom = document.createElement('a');
        pom.setAttribute('href', 'data:text/plain;charset=urf-8,'+encodeURIComponent(text));
        pom.setAttribute('download', filename);
        pom.click();
      };

    // Convert 2d array into comma separated values
    // https://stackoverflow.com/a/14966131
    function convertToCsv(rows) {
        let csvContent = "data:text/csv;charset=utf-8,";

        rows.forEach(function(rowArray) {
            let row = rowArray.join(",");
            csvContent += row + "\r\n";
        });
        return csvContent;
    }
    // Reverse engineered API call to retrieve an array of users in an address list
    // Example AddressListId: "a000a000-0aa0-0a0a-aa00-a000a0000a0a"
    // Example Offset: "300"
    // Example MaxEntriesReturned: "100"
    async function getUsersFromAddressList(BaseFolderId, Offset, MaxEntriesReturned, x_owa_canary) {
        console.log('Performing API request...')
        const response = await fetch("https://outlook.office.com/owa/service.svc?action=FindPeople&app=People", {
            "credentials": "include",
            "headers": {
                "User-Agent": "Mozilla/5.0 (X11; U; Linux x86_64; en-US) Gecko/20072401 Firefox/98.0",
                //"Accept": "*/*",
                "Accept-Language": "en-US,en;q=0.5",
                "action": "FindPeople",
                "content-type": "application/json; charset=utf-8",
                //"ms-cv": "",
                "prefer": "exchange.behavior=\"IncludeThirdPartyOnlineMeetingProviders\"",
                "x-owa-canary": x_owa_canary,
                //"x-owa-correlationid": "",
                //"x-owa-sessionid": "",
                // For a decoded version of x-owa-urlpostdata, please see the bottom of this file
                "x-owa-urlpostdata": "%7B%22__type%22%3A%22FindPeopleJsonRequest%3A%23Exchange%22%2C%22Header%22%3A%7B%22__type%22%3A%22JsonRequestHeaders%3A%23Exchange%22%2C%22RequestServerVersion%22%3A%22V2018_01_08%22%2C%22TimeZoneContext%22%3A%7B%22__type%22%3A%22TimeZoneContext%3A%23Exchange%22%2C%22TimeZoneDefinition%22%3A%7B%22__type%22%3A%22TimeZoneDefinitionType%3A%23Exchange%22%2C%22Id%22%3A%22GMT%20Standard%20Time%22%7D%7D%7D%2C%22Body%22%3A%7B%22IndexedPageItemView%22%3A%7B%22__type%22%3A%22IndexedPageView%3A%23Exchange%22%2C%22BasePoint%22%3A%22Beginning%22%2C%22Offset%22%3A"+Offset+"%2C%22MaxEntriesReturned%22%3A"+MaxEntriesReturned+"%7D%2C%22QueryString%22%3Anull%2C%22ParentFolderId%22%3A%7B%22__type%22%3A%22TargetFolderId%3A%23Exchange%22%2C%22BaseFolderId%22%3A%7B%22__type%22%3A%22AddressListId%3A%23Exchange%22%2C%22Id%22%3A%22"+BaseFolderId+"%22%7D%7D%2C%22PersonaShape%22%3A%7B%22__type%22%3A%22PersonaResponseShape%3A%23Exchange%22%2C%22BaseShape%22%3A%22Default%22%2C%22AdditionalProperties%22%3A%5B%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22PersonaAttributions%22%7D%2C%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22PersonaTitle%22%7D%2C%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22PersonaOfficeLocations%22%7D%5D%7D%2C%22ShouldResolveOneOffEmailAddress%22%3Afalse%2C%22SearchPeopleSuggestionIndex%22%3Afalse%7D%7D",
                //"x-req-source": "People",
                //"Sec-Fetch-Dest": "empty",
                //"Sec-Fetch-Mode": "cors",
                //"Sec-Fetch-Site": "same-origin",
                //"Sec-GPC": "1",
                "Pragma": "no-cache",
                "Cache-Control": "no-cache"
            },
            "method": "POST",
            //"mode": "cors"
        })
            .then(data => data.json());

        // Array(143) [ {…}, {…}, {…} … ]
        let users = response.Body.ResultSet;
        return users
    }

    // Store all extracted user data
    // [[id1, John Smith, jsmith@example.com], [id2, Foo Bar, fbar@example.com]]
    const user_db = [];

    // Get all users
    const users = await getUsersFromAddressList(
        base_folder_id, "0", "1000", canary)

    // Iterate through all users
    for (let index = 0; index < users.length; index++) {
        console.debug('\nAccessing user at index: ' + index)

        // Extract information from user API data
        let user = users[index];
        let displayname = user.DisplayName;
        let emailaddress = user.EmailAddress.EmailAddress
        let id = user.PersonaId.Id;

        // Compile extracted information into an array
        let userdata = [id, displayname, emailaddress];

        // Save compiled user information
        user_db.push(userdata);
        console.debug('New user: ' + userdata);
    }

    // Print user_db array to console
    console.debug(user_db);

    // Download database as a .csv file
    let user_db_csv = convertToCsv(user_db);
    saveAs(user_db_csv, 'user_db.csv');
    console.log('Downloaded results to user_db.csv!')
})();

// Project inspired by: https://github.com/edubey/browser-console-crawl/blob/master/single-story.js

/*
x-owa-urlpostdata decoded:
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
*/
