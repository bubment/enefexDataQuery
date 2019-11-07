
function loginAsync(host, username, password, callback) {

    var params = {};

    postAsyncHandleData(host + "/mobileLogin/login/login", params, function (err, csrfToken){ 


        if (err) {
            callback({success: false, error: err});
        }
        else {
            if (csrfToken !== "var dummy;") {

                params["Login[csrf_token]"] = csrfToken;
                params["Login[loginname]"] = username;
                params["Login[password]"] = password;
                params["Login[language]"] = "hu";
                params["Login[new_password]"] = "";
                params["Login[new_password_again]"] = "";

                postAsyncHandleData(host + "/mobileLogin/login/login", params, function (err, jsonResult) {

                    if (err) {
                        callback({ success: false, error: err });
                    }
                    else {
                        var parseErr;
                        var parsedResult;
                        try {
                            parsedResult = JSON.parse(jsonResult);
                        }
                        catch (e) {
                            parseErr = e;
                        }

                        if (parseErr) {
                            callback({ success: false, error: parseErr });
                        }
                        else {
                            if (parsedResult.success) {
                                createCookie("enefexHost", host);
                                createCookie("enefexUsername", username);
                                createCookie("enefexPassword", password);
                            }
                            callback(null, parsedResult);
                        }
                    }
                });
            }
            else {
                callback(null, { success:true });
            }
        }
    });
}

function logout() {

    var retval = false;

    try {

        params = {};

        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/mobileLogin/login/logout", params);
        var result = JSON.parse(jsonResult);

        retval = result.success;

        if (result.success) {
            eraseCookie("enefexUsername");
            eraseCookie("enefexPassword");
        }
    }
    catch (ex) {
        console.log(ex);
    }

    return retval;
}

function logoutAsync(callback) {

    var host = readCookie("enefexHost");
    var params = {};
    var onLogout = function (err, jsonResult) {
        if (err) {
            callback({ success: false, error: err });
        }
        else {
            var parseError, result;
            try {
                result = JSON.parse(jsonResult);
            }
            catch (e) {
                parseError = e;
            }
            if (parseError) {
                callback({ success: false, error: parseError });
            }
            else {
                if (result.success) {
                    eraseCookie("enefexUsername");
                    eraseCookie("enefexPassword");
                }
                callback(null, result.success);
            }
        }
    };
    postAsyncHandleData(host + "/mobileLogin/login/logout", params, onLogout);
}

function isLoggedIn() {

    var retval = false;

    try {

        var params = [];

        var host = readCookie("enefexHost");
        
        var csrfToken = post(host + "/mobileLogin/login/login", params);

        if (csrfToken == "var dummy;") {
            
            retval = true;
        }
        else {
            retval = false;
        }
    }
    catch (ex) {
        console.log(ex);
    }

    return retval;
}
