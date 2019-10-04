
var startTime = Date.now();
var aTime = [];
var bTime = [];

function getRuntime(index) {
    var retval = 0;
    try {
        if (index == undefined) {
            index = 0;
        }
        if (aTime[index] == undefined) {
            aTime[index] = startTime;
        }
        if (bTime[index] == undefined) {
            bTime[index] = startTime;
        }
        bTime[index] = Date.now();
        var diff = bTime[index] - aTime[index];
        aTime[index] = bTime[index];
        retval = (diff / 1000);
    }
    catch (ex) {
        console.log(ex);
    }
    return retval;
}

function getFullRuntime() {
    var retval = 0;
    try {
        var tmpTime = Date.now();
        var diff = tmpTime - startTime;
        retval = (diff / 1000);
    }
    catch (ex) {
        console.log(ex);
    }
    return retval;
}

var numberSeparator = " ";
var numberComma = ",";

function getSeparatedNumber(x, d) {
    var retval = "";
    try {
        var parts = [x.toString()];

        if (x.toString().split(".").length > 1) {
            parts = x.toString().split(".");
        }
        else {
            parts = x.toString().split(",");
        }

        if (d == undefined) {
            d = 3;
        }

        if ((d != undefined) && (parts.length > 1)) {
            if (parts[parts.length - 1].length > d) {
                parts[parts.length - 1] = parts[parts.length - 1].substr(0, d);

                while (parts[parts.length - 1].endsWith('0')) {
                    parts[parts.length - 1] = parts[parts.length - 1].substr(0, parts[parts.length - 1].length - 1);
                }
            }
        }

        parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, numberSeparator);

        retval = parts.join(numberComma);
        if (retval.endsWith(numberComma)) {
            retval = retval.substr(0, retval.length - 1);
        }
    }
    catch (ex) {
        log(ex);
    }

    return retval;
}

function post(path, params) {

    var retval = "";

    log("HTTP POST (" + path + ", " + JSON.stringify(params) + ": " + getRuntime(2) + " sec", "#bbb");

    try {
        var data = new FormData();
        for (var key in params) {
            if (params.hasOwnProperty(key)) {
                data.append(key, params[key]);
            }
        }

        var host = readCookie("enefexHost");

        $.ajax({
            ///server script to process data
            xhrFields: { withCredentials: true },
            url: path, //web service
            type: "POST",
            contentType: "text/plain",
            ///Form data
            data: data,
            //headers: { 'X-Alt-Referer': host },
            async: false,
            crossDomain: true,

            complete: function () {
                //on complete event     

            },
            progress: function (evt) {
                //progress event  
                alert(evt);
            },
            ///Ajax events
            beforeSend: function (e) {
            },
            success: function (output, status, xhr) {
                retval = output;
            },
            error: function (e) {
                retval = JSON.stringify({
                    success: false,
                    error: {
                        message: e.statusText
                    }
                });
            },
            ///Form data
            //data: data,
            ///Options to tell JQuery not to process data or worry about content-type
            cache: false,
            contentType: false,
            processData: false
        });
    }
    catch (ex) {
        console.log(ex);
        postInProgress = false;
    }


    return retval;
}

function postAsyncHandleData(path, params, callback) {

    log("HTTP POST (" + path + ", " + JSON.stringify(params) + ": " + getRuntime(2) + " sec", "#bbb");

    try {
        var data = new FormData();
        for (var key in params) {
            if (params.hasOwnProperty(key)) {
                data.append(key, params[key]);
            }
        }

        var host = readCookie("enefexHost");

        $.ajax({
            ///server script to process data
            xhrFields: { withCredentials: true },
            url: path, //web service
            type: "POST",

            //contentType: "text/plain",
            ///Options to tell JQuery not to process data or worry about content-type
            contentType: false,
            ///Form data
            data: data,
            //headers: { 'X-Alt-Referer': host },
            async: true,
            crossDomain: true,

            ///Ajax events

            success: function (output, status, xhr) {
                callback(null, output);
            },
            error: function (e) {
                callback(e);
                /*retval = JSON.stringify({
                    success: false,
                    error: {
                        message: e.statusText
                    }
                });*/
            },
            ///Form data
            //data: data,
            ///Options to tell JQuery not to process data or worry about content-type
            cache: false,
            processData: false
        });
    }
    catch (ex) {
        console.log(ex);
        postInProgress = false;
    }


}

// Akkor kell használni ha csak az adatokat tartalmazó JSON objectet várjuk a szerverlekérdezés eredményéül és azt szeretnénk visszaadni a meghívó függvénynek.
function postAsyncGetData(path, params, callback) {

    log("HTTP POST (" + path + ", " + JSON.stringify(params) + ": " + getRuntime(2) + " sec", "#bbb");

    try {
        var data = new FormData();
        for (var key in params) {
            if (params.hasOwnProperty(key)) {
                data.append(key, params[key]);
            }
        }

        var host = readCookie("enefexHost");

        $.ajax({
            ///server script to process data
            xhrFields: { withCredentials: true },
            url: path, //web service
            type: "POST",

            //contentType: "text/plain",
            ///Options to tell JQuery not to process data or worry about content-type
            contentType: false,
            ///Form data
            data: data,
            //headers: { 'X-Alt-Referer': host },
            async: true,
            crossDomain: true,

            ///Ajax events

            success: function (output, status, xhr) {
                var parseError, result;
                try {
                    result = JSON.parse(output);
                }
                catch (e) {
                    parseError = e;
                }
                if (parseError) {
                    callback({ success: false, error: parseError });
                }
                else {
                    callback(null, result);
                }
            },
            error: function (e) {
                callback(e);
                /*retval = JSON.stringify({
                    success: false,
                    error: {
                        message: e.statusText
                    }
                });*/
            },
            ///Form data
            //data: data,
            ///Options to tell JQuery not to process data or worry about content-type
            cache: false,
            processData: false
        });
    }
    catch (ex) {
        console.log(ex);
        postInProgress = false;
    }


}


//************************************************************************
//  Cookie
//************************************************************************

function createCookie(name, value, minutes) {
    try {
        var expires = "";
        if (minutes) {
            var date = new Date();
            date.setTime(date.getTime() + (minutes * 60 * 1000));
            expires = "; expires=" + date.toUTCString();
        }
        document.cookie = encodeURI(name) + "=" + encodeURI(value) + expires + "; path=/";
        //log ("cookie: '" + name +"'='" + value + "'", "#bbb");
    }
    catch (ex) {
        log("Create cookie error (Cookies: " + document.cookie.split(';').length + " cookie, " + getSeparatedNumber(document.cookie.length, 0) + " char)", "#bbb");
        log(ex);
    }
}

function readCookie(name) {
    var nameEQ = encodeURI(name) + "=";
    var ca = document.cookie.split(';');
    for (var i = 0; i < ca.length; i++) {
        var c = ca[i];
        while (c.charAt(0) == ' ') c = c.substring(1, c.length);
        if (c.indexOf(nameEQ) == 0) return decodeURI(c.substring(nameEQ.length, c.length));
    }
    return null;
}

function eraseCookie(name) {
    createCookie(name, "", -1);
}

//************************************************************************
//  Log
//************************************************************************

function log(text, color) {
    if (color != undefined) {
        console.log("%c" + text, "color:" + color + ";");
    }
    else {
        console.log(text);
    }
}
