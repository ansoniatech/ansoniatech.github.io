const labelTemplate = document.getElementById("label0").cloneNode(true); // Clone of initial blank label form for future reference
var lastLabelId = 0; // The number of the last label id

var clientId = "537591206298-9blrp4f9r00s5uo6rat0phlpmijmf25v.apps.googleusercontent.com";
var clientSecret = "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";

var spreadsheetId = "1wVnZ-ikLlnOH0Lq78cNkp30wW3ukTicgZedR_fmJEXA";
var templateSheetId = "982309210";
var currentSheetId = "";
var currentSheetName = "";

function insertBefore(referenceNode, newNode) {

    referenceNode.parentNode.insertBefore(newNode, referenceNode);

}

// Creates a new blank label on the page
function newLabel() {
    if (lastLabelId < 9) {
        const newLabelButton = document.getElementById("new-label");

        var cloneOfLabelTemplate = labelTemplate.cloneNode(true);
        lastLabelId++;
        cloneOfLabelTemplate.id = "label" + lastLabelId;

        insertBefore(newLabelButton, cloneOfLabelTemplate);
    }
}

// Pulls the access token from the url on a redirect and resets the url
function onLoad() {

    var url = window.location;

    accessTokenInURL = new URLSearchParams(url.hash).get('access_token');

    if (accessTokenInURL != null) {
        localStorage.setItem("accessToken", accessTokenInURL);
        window.location = window.location.pathname;
    }

}

// Displays a message for a short period of time
function displayMessage(message) {

    const messageBox = document.getElementById("message-box");

    messageBox.classList.add("message-box-displayed");
    messageBox.innerHTML = message;

    setTimeout(closeMessage, 5 * 1000); // 5 Seconds

}

function closeMessage() {

    const messageBox = document.getElementById("message-box");

    messageBox.classList.remove("message-box-displayed");

}

// Authenticates with Google
function oauthSignIn() {

    // Google's OAuth 2.0 endpoint for requesting an access token
    var oauth2Endpoint = 'https://accounts.google.com/o/oauth2/v2/auth?access_type=offline';
  
    // Create <form> element to submit parameters to OAuth 2.0 endpoint.
    var form = document.createElement('form');
    form.setAttribute('method', 'GET'); // Send as a GET request.
    form.setAttribute('action', oauth2Endpoint);
  
    // Parameters to pass to OAuth 2.0 endpoint.
    // var params = {'client_id': clientId,
    //               'redirect_uri': 'http://127.0.0.1:5500/index.html',
    //               'response_type': 'token',
    //               'scope': 'https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly https://www.googleapis.com/auth/spreadsheets',
    //               'include_granted_scopes': 'true',
    //               'state': 'pass-through value'};
  
    // Parameters to pass to OAuth 2.0 endpoint.
    var params = {'client_id': clientId,
                  'redirect_uri': 'https://ansoniatech.github.io/',
                  'response_type': 'token',
                  'scope': 'https://www.googleapis.com/auth/admin.directory.device.chromeos.readonly https://www.googleapis.com/auth/spreadsheets',
                  'include_granted_scopes': 'true',
                  'state': 'pass-through value'};

    console.log("Sending authentication request to Google:");

    // Add form parameters as hidden input values.
    for (var p in params) {
        var input = document.createElement('input');
        input.setAttribute('type', 'hidden');
        input.setAttribute('name', p);
        input.setAttribute('value', params[p]);
        form.appendChild(input);
        console.log(p + " - " + params[p]);
    }
  
    // Add form to page and submit it to open the OAuth 2.0 endpoint.
    document.body.appendChild(form);
    form.submit();

}

// Populates some of the data of a label using a device's id
function populateById(label, id) {

    console.log("Searching for Chromebook by this ID: " + id);

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);

    // Settings for GET request
    var requestOptions = {
        method: 'GET',
        headers: myHeaders,
        redirect: 'follow'
    };

    var url = "https://admin.googleapis.com/admin/directory/v1/customer/C03ea4clo/devices/chromeos/" + id;

    // Retrieves device info for a chromebook in the console, converts it to JSON, and fills form data
    fetch(url, requestOptions)
    .then(response => response.json())
    .then(result => {
        label.children[0].children[1].children[1].value = result.model;
        label.children[0].children[3].children[1].value = result.annotatedAssetId;
        label.children[0].children[4].children[1].value = result.serialNumber;
        // console.log(result);
        console.log("Success!");
    })
    .catch(error => console.log('error', error));

}

// Retrieves a list of all Chromebooks
function getChromebookList() {

    console.log("Retrieving Chromebook List");

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);

    // Settings for GET request
    var requestOptions = {
        method: 'GET',
        headers: myHeaders,
        redirect: 'follow'
    };

    var url = "https://admin.googleapis.com/admin/directory/v1/customer/C03ea4clo/devices/chromeos/";

    var devices = []; // An array to populate with JSON objects

    // Retrieves a list of all Chrome devices
    fetch(url, requestOptions)
    .then(response => response.json())
    .then(result => {
        console.log(result);
    })
    .catch(error => console.log('error', error));

    console.log(devices.length);

    return devices;

}

// Exports data to the Google Sheet
function exportToSheets() {

    console.log("Exporting to Google Sheets");

    createSheet();

}

// Duplicates the template of the sheet in the Google Sheet
function createSheet() {

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);


    const date = new Date();

    const month = String(date.getMonth() + 1);
    const day = String(date.getDate());
    const year = String(date.getFullYear());

    var dateString = month + "-" + day + "-" + year;

    // var dateString = String(Math.random());

    currentSheetName = dateString;

    // Operations to perform on the spreadsheet
    var requestBody = {
        requests: [
            {
                duplicateSheet: {
                    sourceSheetId: templateSheetId,
                    newSheetName: dateString,
                    insertSheetIndex: 1000
                }
            }
        ]
    }

    // Settings for POST request
    var requestOptions = {
        method: 'POST',
        headers: myHeaders,
        redirect: 'follow',
        includeSpreadSheetInResponse: false,
        body: JSON.stringify(requestBody)
    }

    var url = "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId + ":batchUpdate";

    // Duplicates the template sheet
    fetch(url, requestOptions)
    .then(response => response.json())
    .then(result => {

        if (result.hasOwnProperty("error")) { // Handles errors

            var code = result.error.code;

            if (code == 401) { // Unauthorized
                displayMessage("You are not signed in!");
            } else if (code >= 400) { // Error

                if (result.error.message.includes("already exists")) { // A sheet with the same name already exists
                    displayMessage("You have already created a sheet today! Rename or delete that sheet to create another!");
                } else {
                    displayMessage("An error occurred!");
                }

            }

        } else {

            displayMessage("The sheet was created successfully!");

            currentSheetId = result.replies[0].duplicateSheet.properties.sheetId;

            populateSheet(dateString);

        }

    })
    .catch(error => console.log(error));

}

// Populates the new Google Sheet
function populateSheet(dateString) {

    // Authentication
    var myHeaders = new Headers();
    var authString = "Bearer " + localStorage.getItem("accessToken");
    myHeaders.append("Authorization", authString);

    var ranges = [];

    ranges[0] = currentSheetName + '!B1:B8';
    ranges[1] = currentSheetName + '!D1:D8';
    ranges[2] = currentSheetName + '!B9:B16';
    ranges[3] = currentSheetName + '!D9:D16';
    ranges[4] = currentSheetName + '!B17:B24';
    ranges[5] = currentSheetName + '!D17:D24';
    ranges[6] = currentSheetName + '!B25:B32';
    ranges[7] = currentSheetName + '!D25:D32';
    ranges[8] = currentSheetName + '!B33:B40';
    ranges[9] = currentSheetName + '!D33:D40';
    
    // Changes date to a / format
    dateString = dateString.replace('-', '/');
    dateString = dateString.replace('-', '/');

    var requestData = [];
    var numberOfLabels = lastLabelId + 1;
    for (var i = 0; i < numberOfLabels; i++) {

        var jsonData = {};
        jsonData['majorDimension'] = 'ROWS';
        jsonData['range'] = ranges[i];

        var labelId = "label" + i;
        var selectedLabel = document.getElementById(labelId);
        var inputs = selectedLabel.getElementsByTagName('input');
        var textAreas = selectedLabel.getElementsByTagName('textarea');
        var selection = selectedLabel.getElementsByTagName('select');

        var valueArray = [
            [inputs[0].value],
            [inputs[1].value],
            [textAreas[0].value],
            [inputs[2].value],
            [inputs[3].value],
            [textAreas[1].value],
            [selection[0].options[selection[0].selectedIndex].text],
            [dateString]
        ];

        jsonData['values'] = valueArray;
        
        requestData.push(jsonData);
    }

    // Operations to perform on the spreadsheet
    var requestBody = {
        valueInputOption: 'RAW',
        responseDateTimeRenderOption: 'SERIAL_NUMBER',
        responseValueRenderOption: 'FORMATTED_VALUE',
        includeValuesInResponse: false,
        data: requestData
    }


    // Settings for POST request
    var requestOptions = {
        method: 'POST',
        headers: myHeaders,
        redirect: 'follow',
        includeSpreadSheetInResponse: false,
        body: JSON.stringify(requestBody)
    }

    var url = "https://sheets.googleapis.com/v4/spreadsheets/" + spreadsheetId + "/values:batchUpdate";

    // Duplicates the template sheet
    fetch(url, requestOptions)
    .then(response => response.json())
    .then(result => {

        console.log(result);

        if (result.hasOwnProperty("error")) { // Handles errors

            var code = result.error.code;

            if (code == 401) { // Unauthorized
                displayMessage("You are not signed in!");
            } else if (code >= 400) { // Error
                displayMessage("An error occurred!");
            }

        } else {

            displayMessage("The sheet was populated successfully!");

        }

    })
    .catch(error => console.log(error));

}