// Scrape all users on your Office 365 Outlook people directory
// github.com/smcclennon

// Delete all variables created in this block once execution finishes
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
        department = document.querySelectorAll("[data-log-name=Department]")[0]["textContent"];
        
        return [email, full_name, department];
    }


    // Store all extracted user information
    const all_users = [];

    // Contacts list being displayed 
    contacts = document.getElementsByClassName("ReactVirtualized__Grid__innerScrollContainer")[0]["children"][0];

    // TODO: Iterate through all userids, extracting user information
    // for item in class "ReactVirtualized__Grid__innerScrollContainer"
        // decend to: div role=listitem
        // extract div id: userid = [id=HubPersonaID_...]
            // select this user: clickEvent(userid)

            // Store extracted information for the currently selected user here
            let new_user;

            // Extract currently displayed user
            new_user = getCurrentlyViewedUser();

            // Add new user to the all_users array
            all_users.push(new_user);

            console.log('New user: ' + new_user);
        // }
}

// Inspired by: https://github.com/edubey/browser-console-crawl/blob/master/single-story.js
