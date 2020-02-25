//SZÜKSÉGES GLOBÁLIS VÁLTOZÓK DEFINIÁLÁSA

//Ebbe a változóba menti ki az exceptionHandler függvény a mértékegység nevet ha szükséges
var tmpUnitName;

// Panelek lenyithatóságának engedélyezéséhez szükséges változók
var panelOpenAvailable = true;

// ------------- SEGÉD FÜGGVÉNYEK ELEJE ---------------------------------

// Meghívja az Office.onReady függvény, ami elvileg tesztel, hogy az Add In készen áll-e a futásra.
function officeOnReady() {
    Office.onReady();
}

//Ez a függvény az oldal tetejére görget
function loadAtTop() {
    window.scrollTo(0, 0);
}

//Lenyíló menüt lenyitó és bezáró függvény
function changeDivDisplay(divId, arrowImageID, buttonId) {

    if (panelOpenAvailable) {
        var x = document.getElementById(divId);
        var z = document.getElementById(arrowImageID);
        if (x.style.display !== 'block') {
            x.style.display = 'block';
            z.src = "Images/arrow-down.png";
            location.hash = "#" + buttonId;
        } else {
            x.style.display = 'none';
            z.src = "Images/arrow-left.png";
            loadAtTop();
        }
    }




    
}

//Error message paragrafusokat alaphelyzetbe állító függvény
function setErrorMessageDefault(divId) {
    var y = document.getElementById(divId);
    y.style.display = 'none';
    y.innerHTML = ""
}

// A függvény ami feltölti a lenyíló listákat a szükséges értékekkel
function fillDropDownList() {

    var host = readCookie("enefexHost");

    hours = ["00:00", "00:15", "00:30", "00:45", "01:00", "01:15", "01:30", "01:45", "02:00", "02:15", "02:30", "02:45", "03:00", "03:15", "03:30", "03:45", "04:00", "04:15", "04:30", "04:45", "05:00", "05:15", "05:30", "05:45", "06:00", "06:15", "06:30", "06:45", "07:00", "07:15", "07:30", "07:45", "08:00", "08:15", "08:30", "08:45", "09:00", "09:15", "09:30", "09:45", "10:00", "10:15", "10:30", "10:45", "11:00", "11:15", "11:30", "11:45", "12:00", "12:15", "12:30", "12:45", "13:00", "13:15", "13:30", "13:45", "14:00", "14:15", "14:30", "14:45", "15:00", "15:15", "15:30", "15:45", "16:00", "16:15", "16:30", "16:45", "17:00", "17:15", "17:30", "17:45", "18:00", "18:15", "18:30", "18:45", "19:00", "19:15", "19:30", "19:45", "20:00", "20:15", "20:30", "20:45", "21:00", "21:15", "21:30", "21:45", "22:00", "22:15", "22:30", "22:45", "23:00", "23:15", "23:30", "23:45", "24:00"];


    var meterGroup = function (callback) {
        var fillMeterGroupCallback = function (err, result) {
            if (err) {
                console.log(err.error.message);
                //szükséges helyeken megjeleníteni, hogy a szolgáltatás nem elérhető
            }
            else {
                if (result) {
                    // Mérő csoport listák feltöltése
                    var meterGroupListsArray = [
                        "fogyasztas_osszesito_meter_groups",
                        "feldolgozott_meresek_meter_groups",
                        "heti_jelentes_meter_groups",
                        "szakreferensi_jelentes_meter_groups"
                    ]

                    meterGroupListsArray.forEach(fillMeterGroupLists);

                    var selectedList;
                    var actOption

                    function fillMeterGroupLists(item) {
                        result.forEach(function (element) {
                            selectedList = document.getElementById(item);
                            actOption = document.createElement("option");
                            actOption.text = element.nev;
                            selectedList.add(actOption);
                        });
                    }

                    callback();
                }
                else {
                    var host = readCookie("enefexHost");
                    console.log("A" + host + "/ebill/billing/getMeterGroups lekérdezés tartalma üres");
                }
            }
        }

        params = {};

        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, fillMeterGroupCallback);
    }

    var savedGraphs = function (callback) {
        var fillGetSavedGraphs = function (err, result) {
            if (err) {
                console.log(err.error.message);
                //szükséges helyeken megjeleníteni, hogy a szolgáltatás nem elérhető
            }
            else {
                if (result) {


                    // Mentett beállítások listák feltöltse

                    x = document.getElementById("heti_jelentes_mentett_bealitasok");

                    result.forEach(function (element) {
                        var option = document.createElement("option");
                        option.text = element.name;
                        x.add(option);
                    });

                    callback();
                }
                else {
                    var host = readCookie("enefexHost");
                    console.log("A" + host + "/mdgraph/draw/getSavedGraphs lekérdezés tartalma üres");
                }
            }
        }

        var params = {};

        params["is_public"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/mdgraph/draw/getSavedGraphs", params, fillGetSavedGraphs);

    }
    
    async.parallel(
        [
            meterGroup,
            savedGraphs

        ],
        function (err) {
            console.log('allfinished', err);
        }
    )

    // Heti jelentésnél az órák feltöltése

    x = document.getElementById("heti_jelentes_kezdo_ora");
    y = document.getElementById("heti_jelentes_befejezo_ora");

    hours.forEach(function (element) {
        var option = document.createElement("option");

        if (element == "06:00") {
            option.selected = "selected"
        }
        option.text = element
        x.add(option);
    });

    hours.forEach(function (element) {
        var option = document.createElement("option");

        if (element == "06:00") {
            option.selected = "selected"
        }
        option.text = element
        y.add(option);
    });
}

//A függvény, ami feltölti a szükséges inputokat (Évek, hónapok, stb.) a megfelelő értékekkel
function fillMenuInputs() {
    var actualDate = new Date();
    var currentYear = actualDate.getFullYear();
    var currentMonth = (actualDate.getMonth() + 1);
    var pastMonth = actualDate.getMonth();
    var currentDay = actualDate.getDate();

    if (pastMonth.toString().length == 1) {
        pastMonth = "0" + pastMonth;
    }

    if (currentMonth.toString().length == 1) {
        currentMonth = "0" + currentMonth;
    }
    
    if (currentDay.toString().length == 1) {
        currentDay = "0" + currentDay;
    }

    var dateType1 = currentYear;
    var dateType2 = currentYear + "-" + pastMonth + "-01"
    var dateType3 = currentYear + "-" + currentMonth + "-01"
    var dateType4 = currentYear + "-" + pastMonth

    if (pastMonth == "00") {
        dateType2 = (currentYear - 1) + "-12-01";
        dateType4 = (currentYear - 1) + "-12";
    }


    document.getElementById("szakreferensiJelentesYearFilter").value = dateType1;
    document.getElementById("kezdo_datum").value = dateType2;
    document.getElementById("veg_datum").value = dateType3;
    document.getElementById("onlyYearFilter").value = dateType1;
    document.getElementById("heti_jelentes_kezdo_datum").value = dateType2;
    document.getElementById("heti_jelentes_veg_datum").value = dateType3;
    document.getElementById("szamlaOsszesitoYearFilter").value = dateType1;
}

//Ez a függvény teszteli, hogy a formokban megaott dátumok megdfelelő formátumban lettek-e megadva.
function dateRegExTest(beginDateId, endDateId, errorLabelId) {

    var dateFrom = document.getElementById(beginDateId).value;
    var dateTo = document.getElementById(endDateId).value;
    //Error label definiálása, amibe a Regex ellenőrzések során kiírjuk a hibákat 
    var y = document.getElementById(errorLabelId);

    // Aktuális dátum
    var maxRequestbeginDate = new Date();

    if (Date.parse(dateFrom) > Date.parse(dateTo)) {
        y.style.display = 'block';
        y.innerHTML = "A kezdődátum nagyobb mint a befejező dátum"
        return "RegExTestProblem";
    }

    dateRegexTest = document.getElementById(beginDateId).value
    if (/^\d{4}[\/\-](0?[1-9]|1[012])[\/\-](0?[1-9]|[12][0-9]|3[01])$/
        .test(dateRegexTest) == false) {
        y.style.display = 'block';
        y.innerHTML = "A kezdő dátum nem megfelelő formátumú. Helyes formátum (YYYY-MM-DD)"
        return "RegExTestProblem";
    }

    dateRegexTest = document.getElementById(endDateId).value
    if (/^\d{4}[\/\-](0?[1-9]|1[012])[\/\-](0?[1-9]|[12][0-9]|3[01])$/
        .test(dateRegexTest) == false) {
        y.style.display = 'block';
        y.innerHTML = "A befejező dátum nem megfelelő formátumú. Helyes formátum (YYYY-MM-DD)"
        return "RegExTestProblem";
    }

    if (Date.parse(dateFrom) >= Date.parse(maxRequestbeginDate)) {
        y.style.display = 'block';
        y.innerHTML = "A kezdő dátum túl nagy"
        return "RegExTestProblem";
    }

    //if ((Date.parse(dateTo) - Date.parse(dateFrom)) > 2851200000 ) {
    //    y.style.display = 'block';
    //    y.innerHTML = "A lekérdezés időtartama nagyobb mint 1 hónap"
    //    return "RegExTestProblem";
    //}

    return "RegExTestSuccess";
}

//Ez a függvény a disableElements tömb változóban megadott elementek elérhetőségét (kattinthatóságát és írhatóságát) állítja be
// Ha a status true akkor nem érhetőek el az adott elementek
function changElementsAvailability(disableElements, status) {
    disableElements.forEach(function (element) {
        document.getElementById(element).disabled = status;
    });

    panelOpenAvailable = !status;
}

function setDisableElement(){
    // Azok az elementek (lekérdezés indító gombok) amiket a lekérdezések alatt elérhetetlenné kell tenni, hogy a felhasználó ne tudja módosítani őket
    var importantDisableElements = [];

    var getFuttatoGombok = document.getElementsByClassName("vegleges-futtato-gomb");
    for (var i = 0; i < getFuttatoGombok.length; i++) {
        importantDisableElements.push(getFuttatoGombok[i].id);
    }

    return importantDisableElements;
}

// Ez a függvény kicseréli a lenyíló menük tartalmát egy felmerülő hiba esetén
function placeRequestErrorDiv(replaceableDiv, replaceText) {
    var replacementDiv =
        '<div class="">' +
        '<div class="align-center"><img src="Images/sarga_error.png"></div>' +
        '<div class="align-center requestErrorText">' + replaceText + '</div>' +
        '</div > '

    document.getElementById(replaceableDiv).innerHTML = replacementDiv;
}

//Ellenőrzi, hogy a mainItemInfoArrayban nincs-e duplikált ID
function mainItemInfoPrepareFunction(mainItemInfo) {
    var duplicateTestArray = [];
    var duplicateProblem = false;
    for (var i = 0; i < mainItemInfo.length; i++) {
        for (var j = 0; j < duplicateTestArray.length; j++) {
            if (mainItemInfo[i].mainRequestId.id == duplicateTestArray[j]) {
                console.log("A mainItemInfo tömbben a(z) " + duplicateTestArray[j] + " id többször is benne van");
                duplicateProblem = true;
                break;
            }
        }
        if (duplicateProblem) {
            return "duplicateProblem";
        }

        duplicateTestArray.push(mainItemInfo[i].mainRequestId.id);
    }
    return "TESTOK";
}

// Ez a függvény feltölti az excelDataArrayt a kezdetleges szerkezettel
function excelDataArrayPrepare(mainItemInfo) {
    var excelDataArray = [];
    for (var i = 0; i < mainItemInfo.length; i++) {
        excelDataArray.push(
            {
                "requestId": mainItemInfo[i].mainRequestId.id,
                "sheetName": mainItemInfo[i].sheetName,
                "implementableData": [],

            }
        )
    }

    return excelDataArray;
}

// Ez a függvény a szerverlekérdezésekből jövő adatokat módosítja úgy, ahogyan annak a végén az excelben látszania kell
function exceptionHandler(data, dataProblem) {
    var retval;
    var value;

    if (dataProblem == "mertekegyseg_levagas") {
        value = data
        if (typeof value != "string") {
            console.log("A javítandó változó nem string formátumú!");
            indexOfSpace = -1;
        }
        else {
            indexOfSpace = value.indexOf(" ");
        }
        if (indexOfSpace != -1) {
            retval = value.substr(0, indexOfSpace)
            tmpUnitName = value.substr(indexOfSpace + 1, value.length)
        }
        else {
            retval = "undefined";
            //unitColumnValue = "undefined";
        }

        return retval;
    }

    if (dataProblem == "mertekegyseg_oszlop") {
        if (typeof tmpUnitName != "string") {
            console.log("A kapott mértékegység hibás (nem string)");
            retval = "undefined"
        }
        if (tmpUnitName == "" || tmpUnitName == undefined || tmpUnitName == null) {
            console.log("A kapott mértékegység üres");
            retval = "undefined";
        }
        retval = tmpUnitName;

        return retval;
    }

    if (dataProblem == "specialis_0001") {
        // HHCSID-s csatalkozási pontos problémájának lekezelése
        return "unrepared value";
    }

    return "unreparable value"
}

//Ellenőrzi a paramétermént megadott tömbök esetében, hogy léteznek-e és, hogy nem 0 hosszúságuak-e.
function objectTest(mainItem, mainItemInfo, supportItem) {
    if (
        typeof mainItem == "object" &&
        typeof mainItemInfo == "object" &&
        typeof supportItem == "object" &&
        supportItem.length > 0 &&
        mainItemInfo.length > 0
    ) {
        console.log("Szükséges objectek ellenőrizve");
        return "TESTOK";
    }
    else {
        console.log("A mainItem típusa: " + typeof mainItem);
        console.log("A supportItem típusa: " + typeof supportItem);
        console.log("A result hossza: " + result.length);
        console.log("A supportItem hossza: " + supportItem.length);
        return "Object test problem";
    }
}

// Ezzel a függvénnyel lehet álltani a loader helyzetét. 
// A loaderContainer annak a "panel-loader" osztályú divnek az id-ja, amely az adott panelben lévő loader div-et tartalmazza.
// A loader annak a "loader" osztályú divnek az id-ja ami maga a loader
// A status "block" ha azt akarjuk, hogy előjöjjön a loader, minden más esetbe "none"
function setPanelLoader(loaderContainer, loader, status) {
    if (status == "block") {
        document.getElementById(loaderContainer).style.display = "block"
    }
    else {
        document.getElementById(loaderContainer).style.display = "none"
    }

    var loaderContainerHeight = document.getElementById(loaderContainer).offsetHeight;
    var loaderHeight = document.getElementById(loader).offsetHeight;

    document.getElementById(loader).style.top = ((loaderContainerHeight - loaderHeight) / 2) + "px";
}




// Ha volt már valaki bejelentkezve a programba akkor nem jeleníti meg a felhasznállói instrukciókat.
//!!!!!!!!!!!!!!!!!!!!!!!
//-----------------------------------------------------------------------------------------------

//CSAK AKKOR KELL A wasUserLoggedInTest() FÜGGVÉNY HA AKTIVÁLVA VANNAK AZ INSTRUKCIÓ ÜZENETEK
//function wasUserLoggedInTest() {

//    testval = readCookie("isSomeoneLoggedIn");

//    var instructionMessage = document.getElementById("hideTheBody")

//    if (testval !== null) {
//        instructionMessage.style.display = "none";
//    }
//    else {
//        createCookie("isSomeoneLoggedIn", true)
//    }
//}

//-----------------------------------------------------------------------------------------------
//!!!!!!!!!!!!!!!!!!!!!!!



// ------------- SEGÉD FÜGGVÉNYEK VÉGE ---------------------------------

// ASYNC függvény tesztelés
//Ez lehet jó lesz akkor, ha a teljesítmény javítás érdekében egy nagy függvényt fel kell bontani dátum szerinti több kisebb lekérdezésre;
function asyncSeriesTest() {


    var requiredSheets = ["TEST2", "TEST1"];

    var errorLabel = document.getElementById('fogyasztasOsszesitoError');

    var workSheetHandler = function (callback) {

        var clearableSheet = [];
        var addableSheet = [];

        var separateWorksheets = function (callback) {
            // Ez a függvény a lekérdezéshez szükséges munkalapokat két külön tömbe teszi.
            // A clearableSheet tömbbe teszi a már létező munkalapok nevét
            // Az addableSheet tömbbe teszi a létrehozandó munkalapok nevét
            Excel.run(function (context) {
                var worksheets = context.workbook.worksheets;
                worksheets.load('name');
                return context.sync()
                    .then(function () {
                        var sheetFound;
                        for (var i = 0; i < requiredSheets.length; i++) {
                            sheetFound = false;
                            for (var j = 0; j < worksheets.items.length; j++) {
                                if (requiredSheets[i] == worksheets.items[j].name) {
                                    sheetFound = true;
                                    clearableSheet.push(worksheets.items[j].name);
                                    break;
                                }
                            }
                            if (sheetFound) {
                                continue;
                            }
                            else {
                                addableSheet.push(requiredSheets[i]);
                            }
                        }
                        callback();
                    });
            })

        }

        var clearSheets = function (callback) {
            // A clearSheets függvény tisztítja meg a megadott munkalapok tartalmát
            if (clearableSheet) {
                Excel.run(function (context) {
                    var sheetsNames = clearableSheet;
                    var sheets = context.workbook.worksheets;
                    var sheetsNamesArrayLength = sheetsNames.length;
                    var sheetName;
                    var range;

                    for (var i = 0; i < sheetsNamesArrayLength; i++) {
                        sheetName = sheetsNames[i];
                        range = sheets.getItem(sheetName).getRange();
                        range.load("address");
                        range.clear();

                    }

                    return context.sync()
                        .then(function () {
                            callback();
                        });
                });
            }
        }

        var addSheets = function (callback) {
            // Az addSheets függvény adj hozzá a munkafüzethez a szükséges munkalapokat
            if (addableSheet) {
                Excel.run(function (context) {
                    var newSheets = addableSheet;
                    var sheet = context.workbook.worksheets;
                    var newSheetsArrayLength = newSheets.length;
                    var sheetName;

                    for (var i = 0; i < newSheetsArrayLength; i++) {
                        sheetName = newSheets[i];
                        newSheet = sheet.add(sheetName);
                    }

                    return context.sync()
                        .then(function () {
                            callback();
                        });
                });
            }
        }

        async.series(
            [
                separateWorksheets,
                clearSheets,
                addSheets
            ],
            function (err) {
                console.log('all finished', err);
            }
        );

        callback();
    }

    var asyncMentettBeallitasokGrafikonAdatokElso = function (callback) {

        //var errorLabel = document.getElementById('fogyasztasOsszesitoError');
        errorLabel.style.display = "block";
        errorLabel.innerHTML = "Async Test";
        //y.innerHTML = '<img src="Images/server_messeage_test.gif">'


        //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.

        var getGraphSeriesCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
            }
            else {
                if (result) {


                    //Definiálva vannak az ezen a szinten szükséges változók
                    var headerArray = [];
                    var requiredServerDataArray = [{ dataTag: "tstamp", columnName: "A", headerText: "Dátum" }];
                    var dataLength;
                    var dataInnerLength;


                    var extraInfoObj = result.extraInfo;
                    var extraInfoKeysArray = [];
                    for (var k in extraInfoObj) extraInfoKeysArray.push(k.replace("value", ""));


                    var getMeterTreeCallback = function (err, getMeterTreeResult) {
                        if (err) {
                            document.getElementById('fogyasztasOsszesitoError').innerHTML = err.error.message;
                        }
                        else {
                            if (getMeterTreeResult) {
                                dataLength = Object.keys(getMeterTreeResult.data).length

                                extraInfoKeysArray.forEach(function (element) {
                                    var dataSecondLevelLength;
                                    var found;
                                    for (var i = 0; i < dataLength; i++) {
                                        dataSecondLevelLength = Object.keys(getMeterTreeResult.data[i].data).length
                                        found = false;
                                        for (var j = 0; j < dataSecondLevelLength; j++) {
                                            if (getMeterTreeResult.data[i].data[j].meter_id == element) {

                                                var elementHeaderCompatibleString = "value" + element;
                                                headerArray.push({
                                                    extraInfoKey: elementHeaderCompatibleString,
                                                    extraInfoText: getMeterTreeResult.data[i].data[j].text
                                                });
                                                found = true;
                                                break;
                                            }
                                        }
                                        if (found) {
                                            break;
                                        }
                                    }
                                });



                                for (var i = 0; i < headerArray.length; i++) {
                                    requiredServerDataArray.push({
                                        dataTag: headerArray[i].extraInfoKey,
                                        columnName: excelColumNames[i + 1],
                                        headerText: headerArray[i].extraInfoText
                                    });
                                }

                                dataLength = Object.keys(result.data).length;
                                dataInnerLength = Object.keys(result.data[0]).length;

                                var jsonDataArray = [];
                                var jsonDataInnerArray = [];
                                var d; // result tömben lévő Dátum
                                var correctDateWithFormat; // d változóban lévő dátum YYYY.MM.DD HH:II:SS formátumra hozva 
                                for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                                    jsonDataInnerArray = [];
                                    for (var i = 0; i < requiredServerDataArray.length; i++) {
                                        //result.data[tmpRow].tstamp = correctDateWithFormat;
                                        if (requiredServerDataArray[i].dataTag == "tstamp") {
                                            d = new Date(result.data[tmpRow].tstamp);
                                            correctDateWithFormat = d.getFullYear().toString() + "." + ((d.getMonth() + 1).toString().length == 2 ? (d.getMonth() + 1).toString() : "0" + (d.getMonth() + 1).toString()) + "." + (d.getDate().toString().length == 2 ? d.getDate().toString() : "0" + d.getDate().toString()) + " " + (d.getHours().toString().length == 2 ? d.getHours().toString() : "0" + d.getHours().toString()) + ":" + ((parseInt(d.getMinutes() / 5) * 5).toString().length == 2 ? (parseInt(d.getMinutes() / 5) * 5).toString() : "0" + (parseInt(d.getMinutes() / 5) * 5).toString()) + ":00";

                                            jsonDataInnerArray.push(correctDateWithFormat);
                                        }
                                        else {
                                            jsonDataInnerArray.push(result.data[tmpRow][requiredServerDataArray[i].dataTag]);
                                        }
                                    }
                                    jsonDataArray.push(jsonDataInnerArray);
                                }

                                // ---------------------EXCEL RÉSZ ELEJE --------------------

                                Excel.run(function (context) {

                                    var sheet = context.workbook.worksheets.getItem("TEST1");


                                    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                                    // Excel fejléc kitöltése
                                    for (var i = 0; i < requiredServerDataArray.length; i++) {
                                        sheet.getRange(requiredServerDataArray[i].columnName + "1").values = requiredServerDataArray[i].headerText;
                                        //sheet.getRange("B4").values = "Sample text";
                                    }

                                    //Excel feltöltése adatokkal
                                    var range = sheet.getRange("A2:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                                    range.values = jsonDataArray;
                                    range.untrack();

                                    // Csak a return után lesznek láthatóak az adatok az excelben
                                    boldRange.format.font.bold = true;
                                    return context.sync();
                                })

                                // ---------------------EXCEL RÉSZ VÉGE --------------------

                                errorLabel.style.display = 'none';
                            }
                            else {
                                document.getElementById('fogyasztasOsszesitoError').innerHTML = "A kapott JSON object hibás vagy üres";
                            }
                        }
                    };
                    //Függvény, aminek paramétere a callback függvényünk
                    var host = readCookie("enefexHost");

                    var params = {};
                    params["node"] = "";
                    params["page"] = "";

                    postAsyncGetData(host + "/mdgraph/draw/getMeterTree", params, getMeterTreeCallback);
                }
                else {
                    document.getElementById('fogyasztasOsszesitoError').innerHTML = "A kapott JSON object hibás vagy üres";

                }
            }
        };




        var host = readCookie("enefexHost");
        var params = {};
        //params["datetime_from"] = beginDate;
        //params["datetime_to"] = endDate;
        params["datetime_from"] = "2016-01-01;06:00";
        params["datetime_to"] = "2019-04-01;06:00";
        params["meter_list"] = "76,79,82,85,98,88,96";
        params["baseline_list"] = "";
        params["type_list"] = "2,2,2,2,2,2,2";
        params["serie_type"] = "11";
        params["resolution"] = "0";
        params["type"] = "1";
        params["sendTo"] = "";
        params["checker"] = "0";
        params["extraInfo"] = "1";
        params["fake"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999999";

        postAsyncGetData(host + "/mdgraph/draw/getGraphSeries", params, getGraphSeriesCallback);
        //Caolan async miatt
        callback();
    }

    async.series(
        [
            workSheetHandler,
            asyncMentettBeallitasokGrafikonAdatokElso,
            //asyncMentettBeallitasokGrafikonAdatokMasodik
        ],
        function (err) {
            console.log('allfinished', err);
        }
    )

}

function myTest() {

    //document.getElementById("fogyasztas-osszesito-panel-loader").style.display = "block"
    //var asd = document.getElementById("fogyasztas-osszesito-panel-loader").offsetHeight;
    //var asd1 = document.getElementById("loader-1").offsetHeight;

    //document.getElementById("loader-1").style.top = ((asd - asd1) / 2) + "px";


    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "block");
    
}








