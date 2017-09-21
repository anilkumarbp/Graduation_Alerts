'use strict';

// CONSTANTS - obtained from environment variables
var PORT = process.env.PORT;

// Handle local development and testing
if (process.env.RC_ENVIRONMENT !== 'Production') {
    require('dotenv').config();
}


// Dependencies
var RC = require('ringcentral');
var unirest = require('unirest');
var http = require('http');
var xlsx = require('xlsx');
var server = http.createServer();


// Initialize the sdk for RC
var sdk = new RC({
    server: process.env.RC_API_BASE_URL,
    appKey: process.env.RC_APP_KEY,
    appSecret: process.env.RC_APP_SECRET,
    cachePrefix: process.env.RC_CACHE_PREFIX
});


// Bootstrap Platform and Subscription
var platform = sdk.platform();

var workbook = xlsx.readFile('Master.xlsx');
var sheet_name_list = workbook.SheetNames;
var xlData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]])


//login
init();

/*
 Using Glip Webhook Notifications
 */

function init() {

    /*
     Get the current date
     */
    var today = new Date();
    var dd = today.getDate();
    var mm = today.getMonth()+1;
    var yyyy = today.getFullYear();
    var today = mm + '/' + dd + '/' + yyyy;

    /*
    Create the HTTP JSON Body
     */
    var myObject = {"title": "**Graduation Report for : " + today + "**"};
    var attachements = [];

    /*
    Count of Apps
     */
    var appsGraduated = 0,
        appsDeclined = 0,
        appsPending = 0;


    /*
     Application logic for numbers
     */
    for(var i=0; i < xlData.length; i++) {

        var attachement = {};
        var fields = [];

            if (xlData[i].Grad == "Grad") {
                appsGraduated++;
            }
            if (xlData[i].Grad == "DEC") {
                appsDeclined++;
            }
            if (xlData[i].Grad == "Pending") {
                appsPending++;
            }

        fields.push({"title": "Org Name", "value": xlData[i].OrgName, "short": true});
        fields.push({"title": "App Name", "value": "["+xlData[i].AppName+"]"+"("+xlData[i].AppLink+")", "short": true});
        fields.push({"title": "Scope", "value": xlData[i].Private, "short": true});
        fields.push({"title": "Request Date", "value": xlData[i].RequestDate, "short": true});
        fields.push({"title": "Grad", "value": xlData[i].Grad, "short": true});
        fields.push({"title": "Review Date", "value": xlData[i].ReviewDate, "short": true});
        fields.push({"title": "FreeAcct", "value": xlData[i].FreeAcct, "short": true});

        attachement.fields = fields;
        attachement.color = '#' + ("000000" + Math.random().toString(16).slice(2, 8).toUpperCase()).slice(-6);
        attachements.push(attachement);
    }

    myObject.attachments = attachements;
    myObject.body = "**Graduation Full Report** : https://docs.google.com/spreadsheets/d/1sk-jlI5ZS0vBjnRU_UW7JMSEcz3K9R9HVWozTyseoY0/edit#gid=0" + "\n" + "**Applications Applied for Graduation** : " + xlData.length + "\n" + "**Graduated** : " + appsGraduated + "\n" + "**Declined** : " + appsDeclined + "\n" + "**Pending** : " + appsPending + "\n";
    textProcess(myObject);
}

function textProcess(myObject) {

        return unirest.post(process.env.GLIP_WEBHOOK_URL)
            .headers({'Accept': 'application/json', 'Content-Type': 'application/json'})
            .send(myObject)
            .end(function(response) {
                console.log(response.body);
            });
}
// Start the server
server.listen(PORT);



// Server Event Listeners
server.on('request', inboundRequest);

server.on('error', function (err) {
    console.error(err);
});

server.on('listening', function () {
    console.log('Server is listening to ', PORT);
});

server.on('close', function () {
    console.log('Server has closed and is no longer accepting connections');
});


// Server Request Handler
function inboundRequest(req, res) {
    //console.log('REQUEST: ', req);
}

