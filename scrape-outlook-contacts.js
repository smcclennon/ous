// Scrape all users on your Office 365 Outlook people directory
// github.com/smcclennon

// Delete all variables created in this block once execution finishes
(async () => {
    for (var i = 0; i < 1; i++) {

        // Click elements
        // https://stackoverflow.com/a/22469115
        // Usage: document.getElementById("id").dispatchEvent(clickEvent);
        var clickEvent = new MouseEvent("click", {
            "view": window,
            "bubbles": true,
            "cancelable": false
        });

        // Extract information about the user currently selected/displayed on the webpage
        function getCurrentlyViewedUser() {

            // Variables
            let email;
            let full_name;
            let department;

            // Get email
            email = document.querySelectorAll("[data-log-name=Email]")[1]["children"][0]["children"][0]["children"][0]["children"][0]["children"][1]["textContent"];

            // Get full name
            full_name = document.querySelectorAll("[data-log-name=PersonName]")[0]["textContent"];

            // Get department
            // TODO: Properly wait for the department to load, instead of flooding retry attempts
            let retry = 1000;
            for (i = 0; i < retry; i++) {
                try {
                    department = document.querySelectorAll("[data-log-name=Department]")[0]["textContent"];
                    i = retry;
                } catch (err) {
                    // If the element does not have a department, set this field to '?'
                    department = '?';
                }
            }

            return [email, full_name, department];
        }

        async function extractUserFromID(contactID) {
            const response = await fetch("https://outlook.office.com/owa/service.svc?action=GetPersona&app=People", {
                "credentials": "include",
                "headers": {
                    "User-Agent": "Mozilla/5.0 (X11; U; Linux x86_64; en-US) Gecko/20072401 Firefox/98.0",
                    "Accept": "*/*",
                    "Accept-Language": "en-US,en;q=0.5",
                    "action": "GetPersona",
                    "content-type": "application/json; charset=utf-8",
                    "ms-cv": "xxxxxxxxxx+xxxxxxxxxxx.46",
                    "prefer": "exchange.behavior=\"IncludeThirdPartyOnlineMeetingProviders\"",
                    "x-owa-canary": "-xxxxxxxxxxxxxxxxxxx-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx.",
                    "x-owa-correlationid": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                    "x-owa-sessionid": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
                    "x-owa-urlpostdata": "%7B%22__type%22%3A%22GetPersonaJsonRequest%3A%23Exchange%22%2C%22Header%22%3A%7B%22__type%22%3A%22JsonRequestHeaders%3A%23Exchange%22%2C%22RequestServerVersion%22%3A%22V2018_01_08%22%2C%22TimeZoneContext%22%3A%7B%22__type%22%3A%22TimeZoneContext%3A%23Exchange%22%2C%22TimeZoneDefinition%22%3A%7B%22__type%22%3A%22TimeZoneDefinitionType%3A%23Exchange%22%2C%22Id%22%3A%22GMT%20Standard%20Time%22%7D%7D%7D%2C%22Body%22%3A%7B%22__type%22%3A%22GetPersonaRequest%3A%23Exchange%22%2C%22PersonaId%22%3A%7B%22__type%22%3A%22ItemId%3A%23Exchange%22%2C%22Id%22%3A%22" + contactID + "%22%7D%7D%7D",
                    "x-req-source": "People",
                    "Sec-Fetch-Dest": "empty",
                    "Sec-Fetch-Mode": "cors",
                    "Sec-Fetch-Site": "same-origin",
                    "Sec-GPC": "1",
                    "Pragma": "no-cache",
                    "Cache-Control": "no-cache"
                },
                "method": "POST",
                "mode": "cors"
            })
                .then(data => data.json());
            user = response["Body"]["Persona"];
            emailaddress = user["EmailAddress"]["EmailAddress"];
            displayname = user["DisplayName"];
            //department = ...
            userdata = [emailaddress, displayname, '?'];
            return userdata;
        }

        // Store all extracted user information
        const all_users = [];

        // Contacts list being displayed 
        // HTMLCollection { 0: div, 1: div, 2: div, 3: div, 4: div, 5: div, 6: div, 7: div, 8: div, 9: div, … }
        contacts = document.getElementsByClassName("ReactVirtualized__Grid__innerScrollContainer")[0]["children"];

        // Iterate through all contacts listed
        for (index = 0; index < contacts.length; index++) {
            console.log('iteration loop: ' + index)

            // Obtain contact ID
            let contacts_listitem = contacts[index];
            let contacts_entry = contacts_listitem["children"][0];
            let contacts_entry_id = contacts_entry["id"];
            contacts_entry_id = contacts_entry_id.replace("HubPersonaId_", "");
            console.log('got id: ' + contacts_entry_id);

            // Create variable for storing extracted information
            let new_user;

            // try-catch, as non-users (groups) will not have a department and therefore extraction will fail
            // may also account for contact details taking a while to load
            // Number of times to retry if user info extraction fails
            let retry = 1;
            for (i = 0; i < retry; i++) {
                try {
                    // Extract currently displayed user information and save
                    let user = await extractUserFromID(contacts_entry_id);
                    all_users.push(user);
                    console.log('New user: ' + userdata);

                    // stop retrying, we successfully extracted
                    i = retry;

                } catch (err) {
                    console.log('Error occurred during user info extraction: Attempt ' + i + ': ' + err);
                }
            }
            // TODO: Auto-scroll down the contacts list to load more elements. Wait for elements to load correctly.
        }

        // TODO: Download all_users as a .csv file
        console.log(all_users);
    }
})();
// Inspired by: https://github.com/edubey/browser-console-crawl/blob/master/single-story.js
