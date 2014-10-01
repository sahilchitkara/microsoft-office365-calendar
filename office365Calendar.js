module.exports.Office365Calendar = Office365Calendar;
var util = require('util');
var request = require('request');
function Office365Calendar(accessToken){
    this.accessToken = accessToken;
}

Office365Calendar.prototype.get = function (eventId, done){
    request('https://outlook.office365.com/ews/odata/Me/Events(' + eventId + ')?access_token=' + this.accessToken, function (error, response, body){
        if (error) done(error, null);
        else {
            done(null, JSON.parse(response.body));
        }
    });
};

Office365Calendar.prototype.list = function (done){
    request('https://outlook.office365.com/ews/odata/Me/Events', { 'auth': { 'bearer': this.accessToken } }, function (error, response, body){
        if (error) done(error, null);
        else {
            if (response.body == '')
                done(null, []);
            else
                done(null, JSON.parse(response.body).value);
        }
    });
};

Office365Calendar.prototype.update = function (eventId, body, done){
    request({url: 'https://outlook.office365.com/ews/odata/Me/Events(' + eventId + ')?access_token=' + this.accessToken, method: "PATCH", body: body}, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    });
};

Office365Calendar.prototype.delete = function (eventId, done){
    request.delete('https://outlook.office365.com/ews/odata/Me/Events(' + eventId + ')?access_token=' + this.accessToken, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    });
};

Office365Calendar.prototype.create = function (body, done){
    request({url: 'https://outlook.office365.com/ews/odata/Me/Events?access_token=' + this.accessToken, method: "POST", body: body}, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    })
};

Office365Calendar.prototype.respond = function (eventID, response, body, done){
    request({url: 'https://outlook.office365.com/ews/odata/Me/Events(' + eventID + ')/' + response + '?access_token=' + this.accessToken, method: "POST", body: body}, function (error, response, body){
        if (error) done(error, null);
        else done(null, JSON.parse(response.body));
    })
};
