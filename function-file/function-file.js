/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// The initialize function must be run each time a new page is loaded
Office.initialize = reason => {
    requestToken();
};

function requestToken() {
    $.ajax({
        "async": true,
        "crossDomain": true,
        "url": `https://outlook.office365.com/oauth2/v2.0/token+${Office.context.mailbox.token}`, // Pass your tenant name instead of sharepointtechie
        "method": "POST",
        "headers": {
            "content-type": "application/x-www-form-urlencoded"
        },
        "data": {
            "grant_type": "client_credentials",
            "client_id ": "xxx", //Provide your app id
            "client_secret": "xxx", //Provide your client secret genereated from your app
            "scope ": "https://graph.microsoft.com/.default"
        },
        success: function (response) {
            console.log(response);
            token = response.access_token;
            document.getElementById('content').innerHTML = token;

            $.ajax({
                url: 'https://graph.microsoft.com/v1.0/users/userid/calendarView/delta?startdatetime=2018-12-04T12:11:08Z&enddatetime=2019-01-04T12:11:08Z',
                type: 'GET',
                dataType: 'json',
                beforeSend: function (xhr) {
                    xhr.setRequestHeader('Authorization', 'Bearer '+token+'');
                },
                data: {},
                success: function (results) {                            
                    console.log(response);
                    debugger;
                },
                error: function (error) {
                    console.log("Error in getting data: " + error);
                }
            });
        }

    })
}
// Add any ui-less function here
