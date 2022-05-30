// Scrape users in your Office 365 organisation
// github.com/smcclennon/ous

// Delete all variables created in this block once execution finishes
(async () => {

    // This needs to be replaced with a valid BaseFolderId or the API request will fail
    // How to obtain a BaseFolderId: https://github.com/smcclennon/ous#how-to-get-a-basefolderid
    const base_folder_id = "a000a000-0aa0-0a0a-aa00-a000a0000a0a"

    // Get a cookie from the browser. Used to get the x-owa-canary authentication cookie
    // https://www.tabnine.com/academy/javascript/how-to-get-cookies/
    function getCookie(cName) {
        const name = cName + "=";
        const cDecoded = decodeURIComponent(document.cookie); //to be careful
        const cArr = cDecoded.split('; ');
        let res;
        cArr.forEach(val => {
          if (val.indexOf(name) === 0) res = val.substring(name.length);
        })
        return res
      }

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
        //let csvContent = "data:text/csv;charset=utf-8,";
        let csvContent = "";

        rows.forEach(function(rowArray) {
            let row = rowArray.join(",");
            csvContent += row + "\r\n";
        });
        return csvContent;
    }

    // Reverse engineered API call to retrieve an array of users in an Outlook directory
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
                // For a decoded version of x-owa-urlpostdata, please see: https://github.com/smcclennon/ous#x-owa-urlpostdata-decoded
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

    // Get x-owa-canary
    console.debug('Getting x-owa-canary cookie...')
    const canary = getCookie("X-OWA-CANARY");
    if (canary == undefined) {
		throw "Couldn't retrieve x-owa-canary from your cookies! Please make sure you run this code on a console window for https://outlook.office.com and that you are logged in, then try again."
	} else {
        console.debug('Using x-owa-canary: ' + canary);
    }

    // Store all extracted user data
    // [[id1, John Smith, jsmith@example.com], [id2, Foo Bar, fbar@example.com]]
    const user_db = [];

    // Get all users
    const users = await getUsersFromAddressList(
        base_folder_id, "0", "1000", canary)
        .catch(e => {
            const error_description = "API Request failed. Please check your 'x-owa-canary' is correct and valid.\n\nWe automatically collected this from your cookies, so try logging out and logging back into https://outlook.office.com.\n\ncanary = " + canary + '\n\nAPI request/response error:\n' + e;
            throw error_description;
        }
    );

    console.debug("API request successful!");

    if (users == null | users.length == 0) {
        const error_description = "API Request returned no users. Please check your 'BaseFolderId' is valid. You can find this at the top of the program:\nbase_folder_id = " + base_folder_id + '\n\nHow to obtain a BaseFolderId: https://github.com/smcclennon/ous#how-to-get-a-basefolderid\n\nIt is also possible that the user directory you collected the BaseFolderId for is empty and contains no users. If this is the case, please try using the BaseFolderId for a user directory containing at least 1 user and try again.';
        throw error_description;
    } else {
        console.log('Retrieved ' + users.length + ' users!');
    }

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
    console.debug('\nUser array:')
    console.debug(user_db);

    // Download database as a .csv file
    console.debug('Converting user array to csv...')
    let user_db_csv = convertToCsv(user_db);
    console.debug('Downloading csv...')
    saveAs(user_db_csv, 'user_db.csv');
    console.log('Downloaded results to user_db.csv!')
})();
