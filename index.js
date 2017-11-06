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
    var myObject = {"title": "**Graduation Report as of : " + today + "**"};
    var attachements = [];
    var appLink = "https://ai-developer.ringcentral.com/administration.html#/application/";

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
        fields.push({"title": "App Name", "value": "["+xlData[i].AppName+"]"+"("+ appLink + xlData[i].AppId + ")", "short": true});
        fields.push({"title": "Scope", "value": xlData[i].Scope, "short": true});
        fields.push({"title": "Request Date", "value": xlData[i].RequestDate, "short": true});
        fields.push({"title": "Grad", "value": xlData[i].Grad, "short": true});
        fields.push({"title": "Review Date", "value": xlData[i].ReviewDate, "short": true});
        fields.push({"title": "FreeAcct", "value": xlData[i].FreeAcct, "short": true});

        attachement.fields = fields;
        attachement.color = '#' + (xlData[i].Scope == "Private" ? "0000FF" : "FFA500");
        attachements.push(attachement);
    }

    myObject.attachments = attachements;
    myObject.body = "**Applications Applied for Graduation** : " + xlData.length + "\n" + "**Graduated** : " + appsGraduated + "\n" + "**Declined** : " + appsDeclined + "\n" + "**Pending** : " + appsPending + "\n" + "\n" + "**Daily Graduation Report** : https://docs.google.com/spreadsheets/d/1sk-jlI5ZS0vBjnRU_UW7JMSEcz3K9R9HVWozTyseoY0/edit#gid=0" + "\n" + "\n" + "Here are supplementary graduation weekly and daily sheets with charts for you: https://docs.google.com/a/ringcentral.com/spreadsheets/d/1WVPdmFxFPn-lrwzzEPdacXguDlBT0f_ACzbnMfKnjrU/edit?usp=sharing";
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

