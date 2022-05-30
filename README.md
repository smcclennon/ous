# Office365 User Scraper
Export everyone in your Office 365 organisation into a `.csv` in seconds...

This project reverse engineers the Office 365 Outlook webapp API, collecting all users in an Outlook directory via a single API request.

The `x-owa-canary` cookie is automatically retrieved from your browser and used to authenticate the API request. The API response is then parsed and entered into a 2d array. This array is then converted into a comma-separated-value format which is then downloaded as `user_db.csv` via your browser.

![image](https://user-images.githubusercontent.com/24913281/170899229-8676c592-69e5-4026-9134-61feda1c153f.png)

## Features
- Get full name
- Get email address
- Get Outlook ID
- Export all users to a `.csv` file
- Automatically retrieve the required cookie for API authentication
- Easily extract more information from the API response

## Usage
1. Visit https://outlook.office.com in your browser.
2. Press `F12` to launch the "Developer Tools" popup.
3. Navigate to the "Console" tab within the Developer Tools popup.
4. Paste [this](https://raw.githubusercontent.com/smcclennon/ous/master/scrape-outlook-contacts.js) JavaScript code into your console.
5. Edit the code you just pasted and change `const base_folder_id = ""` so that your BaseFolderId is within the quotation marks. Please see below on how to get a BaseFolderId.
6. Press `Enter` to run the code in the console. Userdata will be printed to the console and downloaded to your computer shortly.

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
10. Copy the `Id` child-key (`["BaseFolderId"]["Id"]`). This should look something like `a000a000-0aa0-0a0a-aa00-a000a0000a0a`.
11. The value you just copied is your BaseFolderId.
