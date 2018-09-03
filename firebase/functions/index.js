'use strict';

const { dialogflow, BasicCard, Button, Suggestions } = require('actions-on-google');
const functions = require('firebase-functions');
var GoogleSpreadsheet = require('google-spreadsheet');
var async = require('async');

const app = dialogflow();

// spreadsheet key is the long id in the sheets URL
var doc = new GoogleSpreadsheet('<spreadsheet key>');
var sheet;

app.intent('INTENT_NAME', (conv) => {

    return googleSheet().then(result => {
        conv.ask(result);
    });

});

exports.dialogflowFirebaseFulfillment = functions.https.onRequest(app);

function googleSheet() {
    return new Promise((resolve, reject) => {
        async.series([
            function setAuth(step) {
                
                var creds = {
                    "private_key": "-----BEGIN PRIVATE KEY-----\n<YOUR_KEY>\n-----END PRIVATE KEY-----\n",
                    "client_email": "<GSERVICE ACCOUNT ID>",
                };
                doc.useServiceAccountAuth(creds, step);
            },
            function getInfoAndWorksheets(step) {
                doc.getInfo(function(err, info) {
                    console.log('Loaded doc: ' + info.title + ' by ' + info.author.email);
                    sheet = info.worksheets[0];
                    console.log('sheet 1: ' + sheet.title + ' ' + sheet.rowCount + 'x' + sheet.colCount);
                    step();
                });
            },
            function workingWithRows(step) {
                // google provides some query options
                sheet.getRows({
                    offset: 1,
                    limit: 20,
                    orderby: 'col2'
                }, function(err, rows) {
                    console.log('Read ' + rows.length + ' rows');
                    step();
                });
            },
            function workingWithCells(step) {
                sheet.getCells({
                    'min-row': 1,
                    'max-row': 15,
                    'return-empty': true
                }, function(err, cells) {
                    for (var i = 0; i < cells.length; i++) {
                        var cell = cells[i];
                        if(cell.value !== '')
                        console.log('Cell R'+cell.row+'C'+cell.col+' = '+cell.value);
                    }
                    var response = '';
                    //form the response according to your need by making logic from the cell value
                    //and add the value to response variable.
                    resolve(response);
                    step();
                });
            }
        ], function(err) {
            if (err) {
                console.log('Error: ' + err);
            }
        });
    });
}