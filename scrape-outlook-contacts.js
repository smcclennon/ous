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
        try {
            department = document.querySelectorAll("[data-log-name=Department]")[0]["textContent"];
        } catch (err) {
            // If the element does not have a department, set this field to '?'
            department = '?';
        }
        
        return [email, full_name, department];
    }


    // Store all extracted user information
    const all_users = [];

    // Contacts list being displayed 
    // HTMLCollection { 0: div, 1: div, 2: div, 3: div, 4: div, 5: div, 6: div, 7: div, 8: div, 9: div, … }
    contacts = document.getElementsByClassName("ReactVirtualized__Grid__innerScrollContainer")[0]["children"];

    // Iterate through all contacts listed
    for (index = 0; index < contacts.length; index++) {

        // Obtain contact element
        let contacts_listitem = contacts[index];
        let contacts_entry = contacts_listitem["children"][0]
        
        // Click contact
        contacts_entry.dispatchEvent(clickEvent);


        // Create variable for storing extracted information
        let new_user;

        // try-catch, as non-users (groups) will not have a department and therefore extraction will fail
        // may also account for contact details taking a while to load
        let retry = 3;
        for (i = 0; i < retry; i++) {
            try {
                // Extract currently displayed user information
                new_user = getCurrentlyViewedUser();

                // Add user information to the all_users array
                all_users.push(new_user);

                console.log('New user: ' + new_user);

                // stop retrying, we successfully extracted
                i = retry;
            
            } catch (err) {
                console.log('Error occurred during user info extraction: Attempt ' + i + ': ' + err);
            }
        }
    }
    // TODO: Download all_users as a .csv file
    console.log(all_users);
}

// Inspired by: https://github.com/edubey/browser-console-crawl/blob/master/single-story.js
