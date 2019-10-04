(function () {
    "use strict";

    
    if (isLoggedIn()) {
       window.location.replace("Home.html");
    }
})();


function pressEnterLogin(event) {
    if (event.keyCode === 13) {
        event.preventDefault();
        onClick_loginButton();
    }

}


/*function onClick_loginButton() {

    var host = $('#host')[0].value;
    var username = $('#username')[0].value;
    var password = $('#password')[0].value;
    var result;

    if (username == "ExcelTest" && password == "Qwert12345@") {
        result = login(host, username, password);

        if (result.success) {
            window.location.replace("alt_Home.html");
            document.getElementById("errorMessage").innerHTML = "";
        }
        else {
            document.getElementById("errorMessage").innerHTML = result.error.message;
        }
    }
    else {
        result = login(host, username, password);

        if (result.success) {
            window.location.replace("Home.html");
            document.getElementById("errorMessage").innerHTML = "";
        }
        else {
            document.getElementById("errorMessage").innerHTML = result.error.message;
        }
    }
}*/

function onClick_loginButton() {

    var host = $('#host')[0].value;
    var username = $('#username')[0].value;
    var password = $('#password')[0].value;


    loginAsync(host, username, password, function (err, result) {

        errorLabel = document.getElementById("errorMessage");

        if (err) {
            errorLabel.innerHTML = err.error.message;
            errorLabel.style.display = "block";
        }
        else {       
            if (result.success) {
                if (username === "ExcelTest" && password === "Qwert12345@") {
                    window.location.replace("alt_Home.html");
                }
                else {
                    window.location.replace("Home.html");
                }
                errorLabel.innerHTML = "";
            }
            else {
                errorLabel.innerHTML = result.error.message;
                errorLabel.style.display = "block";
            }
        }
    }); 
}


function hideTheInstruction() {
    document.getElementById("hideTheBody").style.display = "none";
}

function changeInstruction() {

    var asdddd = document.getElementById("fisrtInstructionText").innerHTML

    if (document.getElementById("fisrtInstructionText").innerHTML == 'Írja be a felhasználónevét és jelszavát az arra kijelölt szövegdobozokba és kattintson a <span class="green-text">ZÖLD</span> Belépés gombra.') {
        document.getElementById("hideTheBody").style.display = "none";
    }
    else {
        document.getElementById("fisrtInstructionText").innerHTML = 'Írja be a felhasználónevét és jelszavát az arra kijelölt szövegdobozokba és kattintson a <span class="green-text">ZÖLD</span> Belépés gombra.'
        document.getElementById("fisrtInstructionImage").src = "Images/login.gif"
    }


}