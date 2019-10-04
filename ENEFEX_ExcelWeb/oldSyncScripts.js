﻿// Server request függvények eleje
function login(host, username, password) {

    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    try {

        var params = {};

        var csrfToken = post(host + "/mobileLogin/login/login", params);

        if (csrfToken != "var dummy;") {
            params = [];

            params["Login[csrf_token]"] = csrfToken;
            params["Login[loginname]"] = username;
            params["Login[password]"] = password;
            params["Login[language]"] = "hu";
            params["Login[new_password]"] = "";
            params["Login[new_password_again]"] = "";

            var jsonResult = post(host + "/mobileLogin/login/login", params);

            retval = JSON.parse(jsonResult);

            if (retval.success) {
                createCookie("enefexHost", host);
                createCookie("enefexUsername", username);
                createCookie("enefexPassword", password);
            }
        }
        else {
            retval = true;
        }
    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;
}

function getFogyasztasOsszesito2(dateFrom, dateTo, meterGroupValue) {
    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    var y = document.getElementById('fogyasztasOsszesitoError');
    y.style.display = 'block';
    y.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>'

    try {
        params = {};

        params["date_from"] = dateFrom;
        params["date_to"] = dateTo;
        params["meter_group"] = meterGroupValue;
        //params["date_from"] = "2019-06-01";
        //params["date_to"] = "2019-07-01";

        params["sendTo"] = "screen";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999999";

        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/ebill/billing/getFogyasztasOsszesito2", params);

        retval = JSON.parse(jsonResult);

    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;

}

function getFeldolgozottMeresek(dateMeasurementStart, meterGroupValue) {
    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    try {
        params = {};

        params["datum_meres_kezdete"] = dateMeasurementStart;
        params["meter_group"] = meterGroupValue;
        params["napok_mutatasa"] = "false";
        params["tankolas_is"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "1000";

        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/ebill/summary/getFeldolgozottMeresek", params);

        retval = JSON.parse(jsonResult);
    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;

}

function getTableStat(notShowAll, dateFrom, dateTo) {
    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    try {
        params = {};

        params["not_show_all"] = notShowAll;
        params["filter_date_interval"] = "1";
        params["date_from"] = dateFrom;
        params["date_to"] = dateTo;
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "25";
        params["sort"] = '[{ "property": "T.id", "direction": "desc" }]';


        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/vstat/baseline/getTableStat", params);

        retval = JSON.parse(jsonResult);

    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;

}

function getMeterGroups() {
    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    try {
        params = {};

        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "25";


        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/ebill/billing/getMeterGroups", params);

        retval = JSON.parse(jsonResult);

    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;

}

function getSavedGraphs() {
    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    try {


        params = {};

        //params["is_public"] = "1";
        //params["page"] = "1";
        //params["start"] = "0";
        //params["limit"] = "25";
        //params["group"] = '[{ "property": "user_name", "direction": "ASC" }]';
        //params["sort"] = '[{ "property": "user_name", "direction": "ASC" }]';

        params["is_public"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";


        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/mdgraph/draw/getSavedGraphs", params);

        retval = JSON.parse(jsonResult);

    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;

}


function getMeterTree() {
    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    try {
        params = {};

        params["node"] = "";
        params["page"] = "";


        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/mdgraph/draw/getMeterTree", params);

        retval = JSON.parse(jsonResult);

    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;
}

function getGraphSeries(dateTimeFrom, dateTimeTo, meterList, typeList, resolution, type) {
    var retval = {
        success: false,
        error: {
            message: ""
        }
    };

    try {
        params = {};

        params["datetime_from"] = dateTimeFrom;
        params["datetime_to"] = dateTimeTo;
        params["meter_list"] = meterList;
        params["baseline_list"] = "";
        params["type_list"] = typeList;
        params["serie_type"] = "11";
        params["resolution"] = resolution;
        params["type"] = type;
        params["sendTo"] = "";
        params["checker"] = "0";
        params["extraInfo"] = "1";
        params["fake"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999999";


        var host = readCookie("enefexHost");

        var jsonResult = post(host + "/mdgraph/draw/getGraphSeries", params);

        retval = JSON.parse(jsonResult);

    }
    catch (ex) {
        console.log(ex);

        retval = {
            success: false,
            error: ex
        };
    }

    return retval;

}



function getMeterTreeAsync(callback) {

    var host = readCookie("enefexHost");
    var params = {};
    params["node"] = "";
    params["page"] = "";

    var getMeterTree = function (err, jsonResult) {
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
                callback(null, result);
            }
        }
    };
    postAsyncHandleData(host + "/mdgraph/draw/getMeterTree", params, getMeterTree);

}

function getGraphSeriesAsync(callback, beginDate, endDate) {

    var host = readCookie("enefexHost");
    var params = {};
    params["datetime_from"] = beginDate;
    params["datetime_to"] = endDate;
    //params["datetime_from"] = "2016-01-01;06:00";
    //params["datetime_to"] = "2019-02-01;06:00";
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

    var getGraphSeries = function (err, jsonResult) {
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
                callback(null, result);
            }
        }
    };
    postAsyncHandleData(host + "/mdgraph/draw/getGraphSeries", params, getGraphSeries);
}

// Server request függvények vége

// Home.js-ben lévő összegző függvények eleje

function onClick_fogyasztasOsszesitoButton() {

    var dateFrom = document.getElementById('kezdo_datum').value;
    var dateTo = document.getElementById('veg_datum').value;

    var meterGroupList = document.getElementById('fogyasztas_osszesito_meter_groups');
    var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;

    var meterGroupArray = getMeterGroups();

    for (var i = 0; i < meterGroupArray.length; i++) {
        if (meterGroupArray[i].nev == meterGroupListSelectedText) {
            var meterGroupValue = meterGroupArray[i].id;
            break;
        }
    }

    var maxRequestbeginDate = new Date();
    // A maxRequestbeginDate változót beállítja, hogy az aktuális Dátum +1 hónap legyen
    //maxRequestbeginDate.setMonth(maxRequestbeginDate.getMonth() + 1);

    var y = document.getElementById('fogyasztasOsszesitoError');

    //Dátumok RegEx validációi
    if (dateRegExTest('kezdo_datum', 'veg_datum', 'fogyasztasOsszesitoError') == "RegExTestProblem") {
        return;
    }

    //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.
    generateWorksheet("IN_F0");
    clearEntireSheet("IN_F0");


    // result egy JSON-t tartalmaz
    var result = getFogyasztasOsszesito2(dateFrom, dateTo, meterGroupValue);

    //Szerver lekérdezés után visszaállítja az error paragrafust láthatatlanná
    y.style.display = 'none';

    // A requiredServerDataArray tartalma
    // dataTag: A JSON 'data' tömbjében levő elemNEK szükséges attribútumát tartalmazza
    // columnName: Azt az Excel oszlop nevet tartalmazza, amibe az adott dataTag kerülni fog
    var requiredServerDataArray = [
        { dataTag: "kategoria_nev", columnName: "A" },
        { dataTag: "epulet_azonosito", columnName: "B" },
        { dataTag: "meter_name", columnName: "C" },
        { dataTag: "meter_identifier", columnName: "D" },
        { dataTag: "pod_azonosito", columnName: "E" },
        { dataTag: "idoszak_kezdete", columnName: "F" },
        { dataTag: "idoszak_vege", columnName: "G" },
        { dataTag: "meres_tipus", columnName: "H" },
        { dataTag: "tarifa_hosszu_nev", columnName: "I" },
        { dataTag: "lekotott_teljesitmeny", columnName: "J" },
        { dataTag: "lekotott_teljesitmeny_mertekegyseg", columnName: "K" },
        { dataTag: "operativ_teljesitmeny", columnName: "L" },
        { dataTag: "operativ_teljesitmeny_mertekegyseg", columnName: "M" },
        { dataTag: "max_teljesitmeny", columnName: "N" },
        { dataTag: "max_teljesitmeny_mertekegyseg", columnName: "O" },
        { dataTag: "fogyasztas", columnName: "P" },
        { dataTag: "fogyasztas_mertekegyseg", columnName: "Q" },
        { dataTag: "fogyasztas_elozo_ev", columnName: "R" },
        { dataTag: "fogyasztas_elozo_ev_mertekegyseg", columnName: "S" },
        { dataTag: "havi_dij", columnName: "T" },
        { dataTag: "havi_dij_mertekegyseg", columnName: "U" },
        { dataTag: "induktiv_tulfogyasztas", columnName: "V" },
        { dataTag: "induktiv_tulfogyasztas_mertekegyseg", columnName: "W" },
        { dataTag: "kapacitiv_fogyasztas", columnName: "X" },
        { dataTag: "kapacitiv_fogyasztas_mertekegyseg", columnName: "Y" },

    ];

    // JSON-ben lévő 'data' tömb hossza
    var dataLength = Object.keys(result.data).length

    // ---------------------EXCEL RÉSZ ELEJE --------------------

    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("IN_F0");

        // Excel fejléc kitöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            sheet.getRange(requiredServerDataArray[i].columnName + "1").values = requiredServerDataArray[i].dataTag;
            //sheet.getRange("B4").values = "Sample text";
        }

        // Excel tartalom feltöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {

                sheet.getRange(requiredServerDataArray[i].columnName + (tmpRow + 2)).values = result.data[tmpRow][requiredServerDataArray[i].dataTag]
                // tmpRow--;
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']
            }
        }
        // Csak a return után lesznek láthatóak az adatok az excelben
        return context.sync();
    })

    // ---------------------EXCEL RÉSZ VÉGE --------------------

}

function onClick_feldolgozottMeresekPanelOpen() {

    var meterGroupList = document.getElementById('feldolgozott_meresek_meter_groups');
    var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;

    var meterGroupArray = getMeterGroups();

    for (var i = 0; i < meterGroupArray.length; i++) {
        if (meterGroupArray[i].nev == meterGroupListSelectedText) {
            var meterGroupValue = meterGroupArray[i].id;
            break;
        }
    }

    //Dátum RegEx validációja
    var y = document.getElementById('feldolgozottMeresekError');
    var maxRequestbeginDate = new Date();
    var currentYear = maxRequestbeginDate.getFullYear();

    if (isNaN(document.getElementById('onlyYearFilter').value) == true) {
        y.style.display = 'block';
        y.innerHTML = "A megadott év nem megfelelő formátumú. Megfelelő formátum (YYYY)"
        return;
    }

    if (document.getElementById('onlyYearFilter').value < 2014) {
        y.style.display = 'block';
        y.innerHTML = "A megadott év 2014 előtti."
        return;
    }

    if (document.getElementById('onlyYearFilter').value > currentYear) {
        y.style.display = 'block';
        y.innerHTML = "A megadott év a jövőben van."
        return;
    }

    generateWorksheet("IN_É0");
    clearEntireSheet("IN_É0");

    var dateMeasurementStart = document.getElementById('onlyYearFilter').value + "-01";

    var result = getFeldolgozottMeresek(dateMeasurementStart, meterGroupValue);

    y.style.display = 'none';

    // A requiredServerDataArray tartalma
    // dataTag: A JSON 'data' tömbjében levő elemnek szükséges attribútumát tartalmazza
    // columnName: Azt az Excel oszlop nevet tartalmazza, amibe az adott dataTag kerülni fog
    var requiredServerDataArray = [
        { dataTag: "identifier", columnName: "A" },
        { dataTag: "name", columnName: "B" },
        { dataTag: "ho1", columnName: "C" },
        { dataTag: "ho2", columnName: "D" },
        { dataTag: "ho3", columnName: "E" },
        { dataTag: "ho4", columnName: "F" },
        { dataTag: "ho5", columnName: "G" },
        { dataTag: "ho6", columnName: "H" },
        { dataTag: "ho7", columnName: "I" },
        { dataTag: "ho8", columnName: "J" },
        { dataTag: "ho9", columnName: "K" },
        { dataTag: "ho10", columnName: "L" },
        { dataTag: "ho11", columnName: "M" },
        { dataTag: "ho12", columnName: "N" },
    ];

    // JSON-ben lévő 'data' tömb hossza
    var dataLength = Object.keys(result.data).length

    // ---------------------EXCEL RÉSZ ELEJE --------------------

    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("IN_É0");

        // Excel fejléc kitöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            sheet.getRange(requiredServerDataArray[i].columnName + "1").values = requiredServerDataArray[i].dataTag;
            //sheet.getRange("B4").values = "Sample text";
        }

        // Excel tartalom feltöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {

                sheet.getRange(requiredServerDataArray[i].columnName + (tmpRow + 2)).values = result.data[tmpRow][requiredServerDataArray[i].dataTag]
                // tmpRow--; --> A VÉGLEGES TESZTNÉL LEHET, HOGY BELE KELL RAKNI EZT A SORT !!!!
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']

            }
        }
        // Csak a return után ltesznek láthatóak az adatok az excelben
        return context.sync();
    })

    // ---------------------EXCEL RÉSZ VÉGE --------------------
}

function hetiFogyasztasOsszesito() {
    var dateFrom = document.getElementById('heti_jelentes_kezdo_datum').value;
    var dateTo = document.getElementById('heti_jelentes_veg_datum').value;

    //Kiválasztott Mérő csoport szövegéhez tartozó ID kinyerése
    var meterGroupList = document.getElementById('heti_jelentes_meter_groups');
    var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;

    var meterGroupArray = getMeterGroups();

    for (var i = 0; i < meterGroupArray.length; i++) {
        if (meterGroupArray[i].nev == meterGroupListSelectedText) {
            var meterGroupValue = meterGroupArray[i].id;
            break;
        }
    }

    // Aktuális dátum
    var maxRequestbeginDate = new Date();
    //Error label definiálása, amibe a Regex ellenőrzések során kiírjuk a hibákat 
    var y = document.getElementById('hetiJelentesError');

    //Dátumok RegEx validációi
    if (dateRegExTest('heti_jelentes_kezdo_datum', 'heti_jelentes_veg_datum', 'hetiJelentesError') == "RegExTestProblem") {
        return;
    }

    //IN_Fö MUNKALAP ELEJE

    //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.
    generateWorksheet("IN_FÖ");
    clearEntireSheet("IN_FÖ");

    // result egy JSON-t tartalmaz
    var result = getFogyasztasOsszesito2(dateFrom, dateTo, meterGroupValue);

    // A requiredServerDataArray tartalma
    // dataTag: A JSON 'data' tömbjében levő elemNEK szükséges attribútumát tartalmazza
    // columnName: Azt az Excel oszlop nevet tartalmazza, amibe az adott dataTag kerülni fog
    var requiredServerDataArray = [
        { dataTag: "kategoria_nev", columnName: "A", headerText: "Kat." },
        { dataTag: "epulet_azonosito", columnName: "B", headerText: "Épület azonositó" },
        { dataTag: "meter_name", columnName: "C", headerText: "Mérés neve" },
        { dataTag: "meter_identifier", columnName: "D", headerText: "Mérő azonosító" },
        { dataTag: "pod_azonosito", columnName: "E", headerText: "POD" },
        { dataTag: "idoszak_kezdete", columnName: "F", headerText: "Időszak kezdete" },
        { dataTag: "idoszak_vege", columnName: "G", headerText: "Időszak vége" },
        { dataTag: "meres_tipus", columnName: "H", headerText: "Mérés típus" },
        { dataTag: "tarifa_hosszu_nev", columnName: "I", headerText: "Tarifa" },
        { dataTag: "lekotott_teljesitmeny", columnName: "J", headerText: "Lekötött telj." },
        { dataTag: "lekotott_teljesitmeny_mertekegyseg", columnName: "K", headerText: "[]" },
        { dataTag: "operativ_teljesitmeny", columnName: "L", headerText: "Operatív teljesítmény" },
        { dataTag: "operativ_teljesitmeny_mertekegyseg", columnName: "M", headerText: "[]" },
        { dataTag: "max_teljesitmeny", columnName: "N", headerText: "Max. telj." },
        { dataTag: "max_teljesitmeny_mertekegyseg", columnName: "O", headerText: "[]" },
        { dataTag: "fogyasztas", columnName: "P", headerText: "Fogyasztás" },
        { dataTag: "fogyasztas_mertekegyseg", columnName: "Q", headerText: "[]" },
        { dataTag: "fogyasztas_elozo_ev", columnName: "R", headerText: "Előző évi fogyasztás" },
        { dataTag: "fogyasztas_elozo_ev_mertekegyseg", columnName: "S", headerText: "[]" },
        { dataTag: "havi_dij", columnName: "T", headerText: "Díj" },
        { dataTag: "havi_dij_mertekegyseg", columnName: "U", headerText: "[]" },
        { dataTag: "induktiv_tulfogyasztas", columnName: "V", headerText: "Induktív túl fogy." },
        { dataTag: "induktiv_tulfogyasztas_mertekegyseg", columnName: "W", headerText: "[]" },
        { dataTag: "kapacitiv_fogyasztas", columnName: "X", headerText: "Kapacitív fogy." },
        { dataTag: "kapacitiv_fogyasztas_mertekegyseg", columnName: "Y", headerText: "[]" },


    ];

    // JSON-ben lévő 'data' tömb hossza
    var dataLength = Object.keys(result.data).length

    // ---------------------EXCEL RÉSZ ELEJE --------------------

    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("IN_FÖ");
        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

        // Excel fejléc kitöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            sheet.getRange(requiredServerDataArray[i].columnName + "1").values = requiredServerDataArray[i].headerText;
            //sheet.getRange("B4").values = "Sample text";
        }

        // Excel tartalom feltöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {

                sheet.getRange(requiredServerDataArray[i].columnName + (tmpRow + 2)).values = result.data[tmpRow][requiredServerDataArray[i].dataTag]
                // tmpRow--;
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']
            }
        }
        // Csak a return után ltesznek láthatóak az adatok az excelben
        boldRange.format.font.bold = true;
        return context.sync();

    })

    // ---------------------EXCEL RÉSZ VÉGE --------------------

    // Nem láthatóvá változtatja az Error Labelt
    y.style.display = 'none';

    //IN_Fö MUNKALAP VÉGE---------------------------------------------------------------------------------------------
}

function hetiAlapVonalSzamertekek() {
    //Kezdő és vég dátumok beolvasása
    var dateFrom = document.getElementById('heti_jelentes_kezdo_datum').value;
    var dateTo = document.getElementById('heti_jelentes_veg_datum').value;

    var csakNemMegFeleloSorokCheckBox = document.getElementById("csakNemMegFeleloSorok");
    //A notShowAll változó lesz a getTableStat lekérdezés notShowAll paramétere
    var notShowAll;
    if (csakNemMegFeleloSorokCheckBox.checked == true) {
        notShowAll = "1";
    }
    else {
        notShowAll = "0";
    }

    // Aktuális dátum
    var maxRequestbeginDate = new Date();
    //Error label definiálása, amibe a Regex ellenőrzések során kiírjuk a hibákat 
    var y = document.getElementById('hetiJelentesError');

    //Dátumok RegEx validációi
    if (dateRegExTest('heti_jelentes_kezdo_datum', 'heti_jelentes_veg_datum', 'hetiJelentesError') == "RegExTestProblem") {
        return;
    }

    //IN_SzA MUNKALAP ELEJE-------------------------------------------------------------------------------------------

    //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.
    generateWorksheet("IN_SzA");
    clearEntireSheet("IN_SzA");

    result = getTableStat(notShowAll, dateFrom, dateTo);

    var requiredServerDataArray = [
        { dataTag: "baseline", columnName: "A", headerText: "Alapvonal" },
        { dataTag: "work_phase", columnName: "B", headerText: "Munkafázis" },
        { dataTag: "identifier", columnName: "C", headerText: "Időszak" },
        { dataTag: "date_from_str", columnName: "D", headerText: "'-tól" },
        { dataTag: "date_to_str", columnName: "E", headerText: "'-ig" },
        { dataTag: "lower_limit", columnName: "F", headerText: "Alsó limit" },
        { dataTag: "upper_limit", columnName: "G", headerText: "Felső limit" },
        { dataTag: "base_unit", columnName: "H", headerText: "Mértékegység" },
        { dataTag: "calculated_value", columnName: "I", headerText: "Érték" },
        { dataTag: "calculated_min_value", columnName: "J", headerText: "Min. érték" },
        { dataTag: "calculated_max_value", columnName: "K", headerText: "Max. érték" },
    ];

    // JSON-ben lévő 'data' tömb hossza
    var dataLength = Object.keys(result.data).length

    // ---------------------EXCEL RÉSZ ELEJE --------------------

    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("IN_SzA");
        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

        // Excel fejléc kitöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            sheet.getRange(requiredServerDataArray[i].columnName + "1").values = requiredServerDataArray[i].headerText;
            //sheet.getRange("B4").values = "Sample text";
        }

        // Excel tartalom feltöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {

                sheet.getRange(requiredServerDataArray[i].columnName + (tmpRow + 2)).values = result.data[tmpRow][requiredServerDataArray[i].dataTag]
                // tmpRow--;
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']
            }
        }
        // Csak a return után ltesznek láthatóak az adatok az excelben

        boldRange.format.font.bold = true;
        return context.sync();
    })

    // ---------------------EXCEL RÉSZ VÉGE --------------------

    // Nem láthatóvá változtatja az Error Labelt
    y.style.display = 'none';
    //IN_SzA MUNKALAP VÉGE--------------------------------------------------------------------------------------------
}

function mentettBeallitasokGrafikonAdatok() {
    //REGEX ELEJE

    //Szükséges ellenőrizendő változók kinyerése a honlapról

    //Kezdő és vég dátumok beolvasása
    var dateFrom = document.getElementById('heti_jelentes_kezdo_datum').value;
    var dateTo = document.getElementById('heti_jelentes_veg_datum').value;

    var dateFromHourList = document.getElementById('heti_jelentes_kezdo_ora');
    var dateFromHourSelectedText = dateFromHourList.options[dateFromHourList.selectedIndex].text;

    var dateToHourList = document.getElementById('heti_jelentes_befejezo_ora');
    var dateToHourSelectedText = dateToHourList.options[dateToHourList.selectedIndex].text;

    //A datetime_from változó lesz a getGraphSeries lekérdezés datetime_from paramétere
    var datetime_from = dateFrom + ";" + dateFromHourSelectedText;

    //A datetime_to változó lesz a getGraphSeries lekérdezés datetime_to paramétere
    var datetime_to = dateTo + ";" + dateToHourSelectedText;

    //Kiválasztott mentett beállítás szövegéhez tartozó mérő csoportok ilyen formátumban ("m142,m12,m5100")
    var savedOptionsList = document.getElementById('heti_jelentes_mentett_bealitasok');
    var savedOptionsListSelectedText = savedOptionsList.options[savedOptionsList.selectedIndex].text;

    var savedOptionsArray = getSavedGraphs();

    for (var i = 0; i < savedOptionsArray.length; i++) {
        if (savedOptionsArray[i].name == savedOptionsListSelectedText) {
            var savedOptionsMeters = savedOptionsArray[i].meters;
            //A savedOptionsType változó lesz a getGraphSeries lekérdezés type paramétere
            var savedOptionsType = savedOptionsArray[i].type;
            //A savedOptionsResolution változó lesz a getGraphSeries lekérdezés resolution paramétere
            var savedOptionsResolution = savedOptionsArray[i].resolution;
            break;
        }
    }
    //A savedOptionsMeters változó lesz a getGraphSeries lekérdezés meter_list paramétere
    savedOptionsMeters = savedOptionsMeters.replace(/m/g, "");

    var savedOptionsMetersArray = new Array();
    savedOptionsMetersArray = savedOptionsMeters.split(",");

    //Innen kezdődik a getGraphSeries lekérdezés type_list paraméterének elkészítése
    var result = getMeterTree();

    var dataLength = Object.keys(result.data).length

    var type_list_string = "";

    savedOptionsMetersArray.forEach(function (element) {

        for (var i = 0; i < dataLength; i++) {
            var dataSecondLevelLength = Object.keys(result.data[i].data).length
            for (var j = 0; j < dataSecondLevelLength; j++) {
                if (result.data[i].data[j].meter_id == element) {

                    type_list_string = type_list_string.concat(result.data[i].data[j].data_type_id, ",");
                    i = dataLength;
                    break;
                }
            }
        }
    });
    //A type_list_string változó lesz getGraphSeries lekérdezés type_list paramétere
    type_list_string = type_list_string.slice(0, -1);

    // Aktuális dátum
    var maxRequestbeginDate = new Date();
    //Error label definiálása, amibe a Regex ellenőrzések során kiírjuk a hibákat 
    var y = document.getElementById('hetiJelentesError');

    //Dátumok RegEx validációi
    if (dateRegExTest('heti_jelentes_kezdo_datum', 'heti_jelentes_veg_datum', 'hetiJelentesError') == "RegExTestProblem") {
        return;
    }

    //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.
    generateWorksheet("IN_FG");
    clearEntireSheet("IN_FG");

    getGraphSeriesResult = getGraphSeries(datetime_from, datetime_to, savedOptionsMeters, type_list_string, savedOptionsResolution, savedOptionsType);

    var extraInfoObj = getGraphSeriesResult.extraInfo;
    var extraInfoKeysArray = [];
    for (var k in extraInfoObj) extraInfoKeysArray.push(k.replace("value", ""));

    var result = getMeterTree();

    var dataLength = Object.keys(result.data).length

    var headerArray = [];
    extraInfoKeysArray.forEach(function (element) {

        for (var i = 0; i < dataLength; i++) {
            var dataSecondLevelLength = Object.keys(result.data[i].data).length
            for (var j = 0; j < dataSecondLevelLength; j++) {
                if (result.data[i].data[j].meter_id == element) {

                    var elementHeaderCompatibleString = "value" + element;
                    headerArray.push({ extraInfoKey: elementHeaderCompatibleString, extraInfoText: result.data[i].data[j].text });
                    i = dataLength;
                    break;
                }
            }
        }
    });

    var requiredServerDataArray = [{ dataTag: "tstamp", columnName: "A", headerText: "Dátum" }];

    var tmp = 1;
    headerArray.forEach(function (element) {
        requiredServerDataArray.push({ dataTag: element.extraInfoKey, columnName: excelColumNames[tmp], headerText: element.extraInfoText });
        tmp++;
    });

    //result = getGraphSeries(datetime_from, datetime_to, savedOptionsMeters, type_list_string, savedOptionsResolution, savedOptionsType);

    // JSON-ben lévő 'data' tömb hossza
    var dataLength = Object.keys(getGraphSeriesResult.data).length

    // ---------------------EXCEL RÉSZ ELEJE --------------------

    var actualLastValueCell;

    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem("IN_FG");
        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

        // Excel fejléc kitöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            sheet.getRange(requiredServerDataArray[i].columnName + "1").values = requiredServerDataArray[i].headerText;
            //sheet.getRange("B4").values = "Sample text";
        }



        // Excel tartalom feltöltése
        for (var i = 0; i < requiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                if (requiredServerDataArray[i].dataTag == "tstamp") {
                    var d = new Date(getGraphSeriesResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                    var correctDateWithFormat = d.getFullYear().toString() + "." + ((d.getMonth() + 1).toString().length == 2 ? (d.getMonth() + 1).toString() : "0" + (d.getMonth() + 1).toString()) + "." + (d.getDate().toString().length == 2 ? d.getDate().toString() : "0" + d.getDate().toString()) + " " + (d.getHours().toString().length == 2 ? d.getHours().toString() : "0" + d.getHours().toString()) + ":" + ((parseInt(d.getMinutes() / 5) * 5).toString().length == 2 ? (parseInt(d.getMinutes() / 5) * 5).toString() : "0" + (parseInt(d.getMinutes() / 5) * 5).toString()) + ":00";
                    sheet.getRange(requiredServerDataArray[i].columnName + (tmpRow + 2)).values = correctDateWithFormat;
                }
                else {
                    sheet.getRange(requiredServerDataArray[i].columnName + (tmpRow + 2)).values = getGraphSeriesResult.data[tmpRow][requiredServerDataArray[i].dataTag]
                    actualLastValueCell = requiredServerDataArray[i].columnName + (tmpRow + 2)
                }
                // tmpRow--;
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']
            }
        }
        // Csak a return után lesznek láthatóak az adatok az excelben
        boldRange.format.font.bold = true;
        return context.sync();

    })

    //var runPrintAllMacroCheckBox = document.getElementById("runPrintAllMacro");

    //if (runPrintAllMacroCheckBox.checked == true) {
    //    setTimeout(function () {
    //        ////Excel.run(function (context) {
    //        ////    var sheet = context.workbook.worksheets.getItem('IN_FG');
    //        ////    var range = sheet.getRange(actualLastValueCell);
    //        ////    range.load("formulas");
    //        ////    return context.sync()
    //        ////        .then(function () {
    //        ////            var rangeData = JSON.stringify(range.formulas);
    //        ////            var jsonResult = JSON.parse(rangeData)

    //        ////            //Az adott rangeből hanyadik sorból és oszlopból vagyunk kiváncsiak az adatra
    //        ////            var myCellValue = jsonResult[0][0];

    //        ////            if (myCellValue != "") {
    //        ////                runPrintAllMacro();
    //        ////                return;
    //        ////            }
    //        ////        });
    //        ////})
    //    }, 5000);
    //}



    // ---------------------EXCEL RÉSZ VÉGE --------------------

    // Nem láthatóvá változtatja az Error Labelt
    y.style.display = 'none';
}

function hetiJelentesKeszito() {

    //var runPrintAllMacroCheckBox = document.getElementById("runPrintAllMacro");

    var queryNumber = 1;

    if (queryNumber == 1) {
        Office.onReady(function () {
            hetiFogyasztasOsszesito();
        });
        queryNumber++;
    }

    if (queryNumber == 2) {
        Office.onReady(function () {
            hetiAlapVonalSzamertekek();
        });
        queryNumber++;
    }

    if (queryNumber == 3) {
        Office.onReady(function () {
            mentettBeallitasokGrafikonAdatok();
        });
        queryNumber++;
    }

    //Office.onReady(function () {
    //    hetiFogyasztasOsszesito();
    //});

    //Office.onReady(function () {
    //    hetiAlapVonalSzamertekek();
    //});

    //Office.onReady(function () {
    //    mentettBeallitasokGrafikonAdatok();
    //});





    //Office.onReady(function () {
    //    if (runPrintAllMacroCheckBox.checked == true) {
    //        runPrintAllMacro();
    //    }
    //});



    //hetiFogyasztasOsszesito();
    //hetiAlapVonalSzamertekek();
    //mentettBeallitasokGrafikonAdatok();

    // A heti jelentések elkészítésénél az lenne a végső cél, hogy egy függvénybe berakjuk a 3 szükséges alfüggvényt és úgy hívjuk meg.
    // Ez sajnos nem lehetséges, mert ha meghívjuk a hetiJelentesKeszito() függvényt akkor a futás végén kidob minket a Loginba a program.
    // Feltehetőleg amiatt, mert a lekérdezések nem szinkron történnek meg és az Office.js API addig nem tudja elérni az Excel kezelőfelületeit, amíg fut a program.
    // Ideiglenesen úgy lett megoldva, hogy a hetiFogyasztasOsszesito() és a hetiAlapVonalSzamertekek() Meghívódnak az onMouseDownra és
    // és a mentettBeallitasokGrafikonAdatok() pedig meghívódik onMouseClickre. és mivel onMouseUp és onMouseClick között frissül a képernyő
    // ezért nem dob a végén hibát.
    // A hibát a lekérdezésekben kéne keresni (sync/async). 
    //Az excel által dobott hibaüzenet a következő:
    //Hiba: Office.js has not fully loaded.Your app must call "Office.onReady()" as part of it's loading sequence
    //    (or set the "Office.initialize" function).If your app has this functionality, try reloading this page.
    // Esetleges segítség: https://docs.microsoft.com/hu-hu/office/dev/add-ins/develop/understanding-the-javascript-api-for-office
}

function ParhuzamosHetiJelentesKeszito() {

    //REGEX-HEZ ÉS LEKÉRDEZÉSEKHEZ SZÜKSÉGES ADATOK KINYERÉSE A HETI JELENTÉS KÉSZÍTŐ FORMRÓL
    //---------------------------------------------------------------------------------------

    //Kezdő és vég dátumok beolvasása
    var dateFrom = document.getElementById('heti_jelentes_kezdo_datum').value;
    var dateTo = document.getElementById('heti_jelentes_veg_datum').value;

    var dateFromHourList = document.getElementById('heti_jelentes_kezdo_ora');
    var dateFromHourSelectedText = dateFromHourList.options[dateFromHourList.selectedIndex].text;

    var dateToHourList = document.getElementById('heti_jelentes_befejezo_ora');
    var dateToHourSelectedText = dateToHourList.options[dateToHourList.selectedIndex].text;

    //A datetime_from változó lesz a getGraphSeries lekérdezés datetime_from paramétere
    var datetime_from = dateFrom + ";" + dateFromHourSelectedText;

    //A datetime_to változó lesz a getGraphSeries lekérdezés datetime_to paramétere
    var datetime_to = dateTo + ";" + dateToHourSelectedText;

    //Kiválasztott mentett beállítás szövegéhez tartozó mérő csoportok ilyen formátumban ("m142,m12,m5100")
    var savedOptionsList = document.getElementById('heti_jelentes_mentett_bealitasok');
    var savedOptionsListSelectedText = savedOptionsList.options[savedOptionsList.selectedIndex].text;

    var csakNemMegFeleloSorokCheckBox = document.getElementById("csakNemMegFeleloSorok");
    //A notShowAll változó lesz a getTableStat lekérdezés notShowAll paramétere
    var notShowAll;
    if (csakNemMegFeleloSorokCheckBox.checked == true) {
        notShowAll = "1";
    }
    else {
        notShowAll = "0";
    }

    //Kiválasztott Mérő csoport szövegéhez tartozó ID kinyerése
    var meterGroupList = document.getElementById('heti_jelentes_meter_groups');
    var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;

    var meterGroupArray = getMeterGroups();

    var meterGroupValue;
    for (var i = 0; i < meterGroupArray.length; i++) {
        if (meterGroupArray[i].nev == meterGroupListSelectedText) {
            // A meterGroupValue változó lesz a getFogyasztasOsszesito2 lekérdezés meterGroupValue paramétere
            meterGroupValue = meterGroupArray[i].id;
            break;
        }
    }

    //---------------------------------------------------------------------------------------

    // Aktuális dátum
    var maxRequestbeginDate = new Date();

    //Error label definiálása, amibe a Regex ellenőrzések során kiírjuk a hibákat 
    var y = document.getElementById('hetiJelentesError');

    //Dátumok RegEx validációi
    if (dateRegExTest('heti_jelentes_kezdo_datum', 'heti_jelentes_veg_datum', 'hetiJelentesError') == "RegExTestProblem") {
        return;
    }

    //IN_Fö MUNKALAP ADATBEOLVASÁS ELEJE
    //---------------------------------------------------------------------------------------

    //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.
    generateWorksheet("IN_FÖ");
    clearEntireSheet("IN_FÖ");

    // result egy JSON-t tartalmaz
    var getFogyasztasOsszesitoResult = getFogyasztasOsszesito2(dateFrom, dateTo, meterGroupValue);

    // A requiredServerDataArray tartalma
    // dataTag: A JSON 'data' tömbjében levő elemNEK szükséges attribútumát tartalmazza
    // columnName: Azt az Excel oszlop nevet tartalmazza, amibe az adott dataTag kerülni fog
    var getFogyasztasOsszesitoRequiredServerDataArray = [
        { dataTag: "kategoria_nev", columnName: "A", headerText: "Kat." },
        { dataTag: "epulet_azonosito", columnName: "B", headerText: "Épület azonositó" },
        { dataTag: "meter_name", columnName: "C", headerText: "Mérés neve" },
        { dataTag: "meter_identifier", columnName: "D", headerText: "Mérő azonosító" },
        { dataTag: "pod_azonosito", columnName: "E", headerText: "POD" },
        { dataTag: "idoszak_kezdete", columnName: "F", headerText: "Időszak kezdete" },
        { dataTag: "idoszak_vege", columnName: "G", headerText: "Időszak vége" },
        { dataTag: "meres_tipus", columnName: "H", headerText: "Mérés típus" },
        { dataTag: "tarifa_hosszu_nev", columnName: "I", headerText: "Tarifa" },
        { dataTag: "lekotott_teljesitmeny", columnName: "J", headerText: "Lekötött telj." },
        { dataTag: "lekotott_teljesitmeny_mertekegyseg", columnName: "K", headerText: "[]" },
        { dataTag: "operativ_teljesitmeny", columnName: "L", headerText: "Operatív teljesítmény" },
        { dataTag: "operativ_teljesitmeny_mertekegyseg", columnName: "M", headerText: "[]" },
        { dataTag: "max_teljesitmeny", columnName: "N", headerText: "Max. telj." },
        { dataTag: "max_teljesitmeny_mertekegyseg", columnName: "O", headerText: "[]" },
        { dataTag: "fogyasztas", columnName: "P", headerText: "Fogyasztás" },
        { dataTag: "fogyasztas_mertekegyseg", columnName: "Q", headerText: "[]" },
        { dataTag: "fogyasztas_elozo_ev", columnName: "R", headerText: "Előző évi fogyasztás" },
        { dataTag: "fogyasztas_elozo_ev_mertekegyseg", columnName: "S", headerText: "[]" },
        { dataTag: "havi_dij", columnName: "T", headerText: "Díj" },
        { dataTag: "havi_dij_mertekegyseg", columnName: "U", headerText: "[]" },
        { dataTag: "induktiv_tulfogyasztas", columnName: "V", headerText: "Induktív túl fogy." },
        { dataTag: "induktiv_tulfogyasztas_mertekegyseg", columnName: "W", headerText: "[]" },
        { dataTag: "kapacitiv_fogyasztas", columnName: "X", headerText: "Kapacitív fogy." },
        { dataTag: "kapacitiv_fogyasztas_mertekegyseg", columnName: "Y", headerText: "[]" },
    ];



    //---------------------------------------------------------------------------------------
    //IN_Fö MUNKALAP ADATBEOLVASÁS VÉGE

    //IN_SzA MUNKALAP ADATBEOLVASÁS ELEJE
    //---------------------------------------------------------------------------------------

    //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.
    generateWorksheet("IN_SzA");
    clearEntireSheet("IN_SzA");

    getTableStatResult = getTableStat(notShowAll, dateFrom, dateTo);

    var getTableStatRequiredServerDataArray = [
        { dataTag: "baseline", columnName: "A", headerText: "Alapvonal" },
        { dataTag: "work_phase", columnName: "B", headerText: "Munkafázis" },
        { dataTag: "identifier", columnName: "C", headerText: "Időszak" },
        { dataTag: "date_from_str", columnName: "D", headerText: "'-tól" },
        { dataTag: "date_to_str", columnName: "E", headerText: "'-ig" },
        { dataTag: "lower_limit", columnName: "F", headerText: "Alsó limit" },
        { dataTag: "upper_limit", columnName: "G", headerText: "Felső limit" },
        { dataTag: "base_unit", columnName: "H", headerText: "Mértékegység" },
        { dataTag: "calculated_value", columnName: "I", headerText: "Érték" },
        { dataTag: "calculated_min_value", columnName: "J", headerText: "Min. érték" },
        { dataTag: "calculated_max_value", columnName: "K", headerText: "Max. érték" },
    ];

    //---------------------------------------------------------------------------------------
    //IN_SzA MUNKALAP ADATBEOLVASÁS VÉGE

    //IN_FG MUNKALAP ADATBEOLVASÁS ELEJE
    //---------------------------------------------------------------------------------------
    var savedOptionsArray = getSavedGraphs();

    for (var i = 0; i < savedOptionsArray.length; i++) {
        if (savedOptionsArray[i].name == savedOptionsListSelectedText) {
            var savedOptionsMeters = savedOptionsArray[i].meters;
            //A savedOptionsType változó lesz a getGraphSeries lekérdezés type paramétere
            var savedOptionsType = savedOptionsArray[i].type;
            //A savedOptionsResolution változó lesz a getGraphSeries lekérdezés resolution paramétere
            var savedOptionsResolution = savedOptionsArray[i].resolution;
            break;
        }
    }
    //A savedOptionsMeters változó lesz a getGraphSeries lekérdezés meter_list paramétere
    savedOptionsMeters = savedOptionsMeters.replace(/m/g, "");

    var savedOptionsMetersArray = new Array();
    savedOptionsMetersArray = savedOptionsMeters.split(",");

    //Innen kezdődik a getGraphSeries lekérdezés type_list paraméterének elkészítése
    var getMeterTreeResult = getMeterTree();

    var dataLength = Object.keys(getMeterTreeResult.data).length

    var type_list_string = "";

    savedOptionsMetersArray.forEach(function (element) {

        for (var i = 0; i < dataLength; i++) {
            var dataSecondLevelLength = Object.keys(getMeterTreeResult.data[i].data).length
            for (var j = 0; j < dataSecondLevelLength; j++) {
                if (getMeterTreeResult.data[i].data[j].meter_id == element) {

                    type_list_string = type_list_string.concat(getMeterTreeResult.data[i].data[j].data_type_id, ",");
                    i = dataLength;
                    break;
                }
            }
        }
    });
    //A type_list_string változó lesz getGraphSeries lekérdezés type_list paramétere
    type_list_string = type_list_string.slice(0, -1);

    //Legenerálja a szükséges munkalapot, vagy ha létezett akkor aktiválja és kitörli a tartalmát.
    generateWorksheet("IN_FG");
    clearEntireSheet("IN_FG");

    getGraphSeriesResult = getGraphSeries(datetime_from, datetime_to, savedOptionsMeters, type_list_string, savedOptionsResolution, savedOptionsType);

    var extraInfoObj = getGraphSeriesResult.extraInfo;
    var extraInfoKeysArray = [];
    for (var k in extraInfoObj) extraInfoKeysArray.push(k.replace("value", ""));

    var dataLength = Object.keys(getMeterTreeResult.data).length

    var headerArray = [];
    extraInfoKeysArray.forEach(function (element) {

        for (var i = 0; i < dataLength; i++) {
            var dataSecondLevelLength = Object.keys(getMeterTreeResult.data[i].data).length
            for (var j = 0; j < dataSecondLevelLength; j++) {
                if (getMeterTreeResult.data[i].data[j].meter_id == element) {

                    var elementHeaderCompatibleString = "value" + element;
                    headerArray.push({ extraInfoKey: elementHeaderCompatibleString, extraInfoText: getMeterTreeResult.data[i].data[j].text });
                    i = dataLength;
                    break;
                }
            }
        }
    });

    var getGraphSeriesRequiredServerDataArray = [{ dataTag: "tstamp", columnName: "A", headerText: "Dátum" }];

    var tmp = 1;
    headerArray.forEach(function (element) {
        getGraphSeriesRequiredServerDataArray.push({ dataTag: element.extraInfoKey, columnName: excelColumNames[tmp], headerText: element.extraInfoText });
        tmp++;
    });

    //---------------------------------------------------------------------------------------
    //IN_FG MUNKALAP ADATBEOLVASÁS VÉGE

    // ---------------------EXCELBE ADATBEOLVASÁS ELEJE --------------------

    // Az excelbe bemásolandó szerverről lekérdezett JSON-ben lévő data tömb hossza.
    var dataLength;

    Excel.run(function (context) {
        // ---------------------IN_FÖ MUNKALAP BEOLVASÁS ELEJE --------------------
        dataLength = Object.keys(getFogyasztasOsszesitoResult.data).length

        var sheet = context.workbook.worksheets.getItem("IN_FÖ");
        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

        // Excel fejléc kitöltése
        for (var i = 0; i < getFogyasztasOsszesitoRequiredServerDataArray.length; i++) {
            sheet.getRange(getFogyasztasOsszesitoRequiredServerDataArray[i].columnName + "1").values = getFogyasztasOsszesitoRequiredServerDataArray[i].headerText;
            //sheet.getRange("B4").values = "Sample text";
        }

        // Excel tartalom feltöltése
        for (var i = 0; i < getFogyasztasOsszesitoRequiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {

                sheet.getRange(getFogyasztasOsszesitoRequiredServerDataArray[i].columnName + (tmpRow + 2)).values = getFogyasztasOsszesitoResult.data[tmpRow][getFogyasztasOsszesitoRequiredServerDataArray[i].dataTag]
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']
            }
        }
        // Csak a return után ltesznek láthatóak az adatok az excelben
        boldRange.format.font.bold = true;
        // ---------------------IN_FÖ MUNKALAP BEOLVASÁS VÉGE --------------------

        // ---------------------IN_SzA MUNKALAP BEOLVASÁS ELEJE --------------------
        dataLength = Object.keys(getTableStatResult.data).length

        var sheet = context.workbook.worksheets.getItem("IN_SzA");
        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

        // Excel fejléc kitöltése
        for (var i = 0; i < getTableStatRequiredServerDataArray.length; i++) {
            sheet.getRange(getTableStatRequiredServerDataArray[i].columnName + "1").values = getTableStatRequiredServerDataArray[i].headerText;
            //sheet.getRange("B4").values = "Sample text";
        }

        // Excel tartalom feltöltése
        for (var i = 0; i < getTableStatRequiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {

                sheet.getRange(getTableStatRequiredServerDataArray[i].columnName + (tmpRow + 2)).values = getTableStatResult.data[tmpRow][getTableStatRequiredServerDataArray[i].dataTag]
                // tmpRow--;
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']
            }
        }
        // Csak a return után ltesznek láthatóak az adatok az excelben

        boldRange.format.font.bold = true;

        // ---------------------IN_SzA MUNKALAP BEOLVASÁS VÉGE --------------------

        // ---------------------IN_FG MUNKALAP BEOLVASÁS ELEJE --------------------
        dataLength = Object.keys(getGraphSeriesResult.data).length

        var sheet = context.workbook.worksheets.getItem("IN_FG");
        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

        // Excel fejléc kitöltése
        for (var i = 0; i < getGraphSeriesRequiredServerDataArray.length; i++) {
            sheet.getRange(getGraphSeriesRequiredServerDataArray[i].columnName + "1").values = getGraphSeriesRequiredServerDataArray[i].headerText;
            //sheet.getRange("B4").values = "Sample text";
        }



        // Excel tartalom feltöltése
        for (var i = 0; i < getGraphSeriesRequiredServerDataArray.length; i++) {
            for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                if (getGraphSeriesRequiredServerDataArray[i].dataTag == "tstamp") {
                    var d = new Date(getGraphSeriesResult.data[tmpRow][getGraphSeriesRequiredServerDataArray[i].dataTag]);
                    var correctDateWithFormat = d.getFullYear().toString() + "." + ((d.getMonth() + 1).toString().length == 2 ? (d.getMonth() + 1).toString() : "0" + (d.getMonth() + 1).toString()) + "." + (d.getDate().toString().length == 2 ? d.getDate().toString() : "0" + d.getDate().toString()) + " " + (d.getHours().toString().length == 2 ? d.getHours().toString() : "0" + d.getHours().toString()) + ":" + ((parseInt(d.getMinutes() / 5) * 5).toString().length == 2 ? (parseInt(d.getMinutes() / 5) * 5).toString() : "0" + (parseInt(d.getMinutes() / 5) * 5).toString()) + ":00";
                    sheet.getRange(getGraphSeriesRequiredServerDataArray[i].columnName + (tmpRow + 2)).values = correctDateWithFormat;
                }
                else {
                    sheet.getRange(getGraphSeriesRequiredServerDataArray[i].columnName + (tmpRow + 2)).values = getGraphSeriesResult.data[tmpRow][getGraphSeriesRequiredServerDataArray[i].dataTag]
                    actualLastValueCell = getGraphSeriesRequiredServerDataArray[i].columnName + (tmpRow + 2)
                }
                // tmpRow--;
                //Minta: sheet.getRange("B4").values = "Sample text";
                //Minta: result.data['data tömb hanyadik eleme']['választott elem melyik tulajdonsága']
            }
        }
        // Csak a return után lesznek láthatóak az adatok az excelben
        boldRange.format.font.bold = true;

        // ---------------------IN_FG MUNKALAP BEOLVASÁS VÉGE --------------------

        return context.sync();
    })

    // ---------------------EXCELBE ADATBEOLVASÁS VÉGE --------------------
}
// Home.js-ben lévő összegző függvények vége

//supportfunctions.js-ben található feleslegessé vált kódrészletek eleje



// Függvény ami kigenerálja a paraméterként megadott nevű munkalapot, vagy ha az már létezik, akkor csak aktiválja
function generateWorksheet(sheetName) {

    Excel.run(function (context) {
        var sheet = context.workbook.worksheets;
        var newSheet = sheet.add(sheetName);
        sheet.load("name, position");

        //sheet = context.workbook.worksheets.getItem(sheetName);
        //sheet.activate();

        //var activablesheet = context.workbook.worksheets.getItem(sheetName);
        //activablesheet.activate();

        //activablesheet.load("name, position");
        return context.sync()
            .catch(function () {
                console.log("generateWorksheet function has a problem with gerneate " + sheetName + " worksheet.");
            });
    });
}

// Függvény, ami kitörli az aktív munkalap teljes tartalmát
function clearEntireSheet(sheetName) {
    Excel.run(function (context) {
        var sheet = context.workbook.worksheets.getItem(sheetName);
        var range = sheet.getRange();
        range.clear();

        return context.sync()
            .catch(function () {
                console.log("clearEntireSheet function has a problem with clear " + sheetName + " worksheet.");
            });
    })
}

// A függvény ami kiirajta az error paragrafusba, hogy a Szerverlekérdezés folyamatban
function serverIsRunningMessage(errorLabelId) {
    var y = document.getElementById(errorLabelId);

    y.style.display = 'block';
    y.style.visibility = 'visible';
    y.hidden = false;
    y.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>'
    return;
}

// A függvény ami kiirajta az error paragrafusba, hogy a Szerverlekérdezés sikertelen
function serverIsRunningFalseMessage(errorLabelId) {
    var y = document.getElementById(errorLabelId);
    asd = y.innerHTML;
    if (y.innerHTML == '<span class="green-text">Szerverlekérdezés folymatban...</span>') {
        y.innerHTML = '<span class="red-text">Szerverlekérdezés sikertelen</span>'
        return;
    }

}

// Ez a régebbi csak EGY DARAB munkalapot generál
function generateWorksheetCtx(context, sheetName) {
    var sheet = context.workbook.worksheets;
    var newSheet = sheet.add(sheetName);
    newSheet.load("name, position");
}

function clearSheet(context, sheetName) {

    var sheets = context.workbook.worksheets;
    var range;
    range = sheets.getItem(sheetName).getRange();
    range.load("address");
    range.clear();


}

//Ez az újabbik egy TÖMB ALAPJÁN gernerálja ki a munkalapokat
function generateWorkSheetContext(context, newSheets) {
    var sheet = context.workbook.worksheets;
    var newSheetsArrayLength = newSheets.length;
    var sheetName;

    for (var i = 0; i < newSheetsArrayLength; i++) {
        sheetName = newSheets[i];
        newSheet = sheet.add(sheetName);
    }

}

function deleteWorksheetCtx(context, sheetName) {
    var sheet = context.workbook.worksheets.getItem(sheetName);
    sheet.delete();

    return context.sync();
}

function clearEntireSheetCtx(sheetsNames) {
    //var sheet = context.workbook.worksheets.getItem(sheetName);
    //sheet.getRange().clear();

    //sheet.load("name, position");

    Excel.run(function (context) {
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

        return context.sync();
    });
}

// A workSheetHandler függvény a megadott requiredSheets tömbben lévő munkalapok közül a létezőket tisztítja, a nem létezőket létrehozza
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

// A függvény ami feltölti a heti jelentéseknél a Mentett beállítási lehetőségeket
function fillSavedOptionsList() {
    var result = getSavedGraphs();

    var x = document.getElementById("heti_jelentes_mentett_bealitasok");

    result.forEach(function (element) {
        var option = document.createElement("option");
        option.text = element.name;
        x.add(option);
    });
}

// A függvény ami a fogyasztásösszesítőnél kezeli a kezdő és a vég dátumot
function fogyasztasOsszessitoDate() {

    var dateFrom = document.getElementById('kezdo_datum').value;
    var y = document.getElementById('fogyasztasOsszesitoError');
    var maxRequestbeginDate = new Date();

    dateRegexTest = document.getElementById('kezdo_datum').value
    //if (/^\d{4}[\/\-](0?[1-9]|1[012])[\/\-](0?[1-9]|[12][0-9]|3[01])$/
    if (/^[0-9]{4}-(((0[13578]|(10|12))-(0[1-9]|[1-2][0-9]|3[0-1]))|(02-(0[1-9]|[1-2][0-9]))|((0[469]|11)-(0[1-9]|[1-2][0-9]|30)))$/
        .test(dateRegexTest) == false) {
        y.style.display = 'block';
        y.innerHTML = "A kezdő dátum nem megfelelő formátumú. Helyes formátum (YYYY-MM-DD)"
        return;
    }

    if (Date.parse(dateFrom) >= Date.parse(maxRequestbeginDate)) {
        y.style.display = 'block';
        y.innerHTML = "A kezdő dátum túl nagy"
        return;
    }

    var dateTo = new Date(document.getElementById('kezdo_datum').value);
    dateTo.setMonth(dateTo.getMonth() + 1);
    var dateToYear = dateTo.getFullYear();

    var dateToMonth = (dateTo.getMonth() + 1);
    if (dateToMonth < 10) {
        dateToMonth = "0" + dateToMonth;
    }

    var dateToDay = dateTo.getDate();
    if (dateToDay < 10) {
        dateToDay = "0" + dateToDay;
    }

    dateTo = dateToYear + "-" + dateToMonth + "-" + dateToDay;
    document.getElementById('veg_datum').value = dateTo;
}

function fillDropDownList() {

    // Mérő csoport listák feltöltése

    var result = getMeterGroups();
    var x = document.getElementById("fogyasztas_osszesito_meter_groups");
    var y = document.getElementById("feldolgozott_meresek_meter_groups");
    var z = document.getElementById("heti_jelentes_meter_groups");



    result.forEach(function (element) {
        var option = document.createElement("option");
        option.text = element.nev;
        x.add(option);
    });

    result.forEach(function (element) {
        var option = document.createElement("option");
        option.text = element.nev;
        y.add(option);
    });

    result.forEach(function (element) {
        var option = document.createElement("option");
        option.text = element.nev;
        z.add(option);
    });

    // Mentett beállítás listák feltöltése
    result = getSavedGraphs();

    x = document.getElementById("heti_jelentes_mentett_bealitasok");

    result.forEach(function (element) {
        var option = document.createElement("option");
        option.text = element.name;
        x.add(option);
    });

    // Heti jelentésnél az órák feltöltése
    hours = ["00:00", "00:15", "00:30", "00:45", "01:00", "01:15", "01:30", "01:45", "02:00", "02:15", "02:30", "02:45", "03:00", "03:15", "03:30", "03:45", "04:00", "04:15", "04:30", "04:45", "05:00", "05:15", "05:30", "05:45", "06:00", "06:15", "06:30", "06:45", "07:00", "07:15", "07:30", "07:45", "08:00", "08:15", "08:30", "08:45", "09:00", "09:15", "09:30", "09:45", "10:00", "10:15", "10:30", "10:45", "11:00", "11:15", "11:30", "11:45", "12:00", "12:15", "12:30", "12:45", "13:00", "13:15", "13:30", "13:45", "14:00", "14:15", "14:30", "14:45", "15:00", "15:15", "15:30", "15:45", "16:00", "16:15", "16:30", "16:45", "17:00", "17:15", "17:30", "17:45", "18:00", "18:15", "18:30", "18:45", "19:00", "19:15", "19:30", "19:45", "20:00", "20:15", "20:30", "20:45", "21:00", "21:15", "21:30", "21:45", "22:00", "22:15", "22:30", "22:45", "23:00", "23:15", "23:30", "23:45", "24:00"];
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
//supportfunctions.js-ben található feleslegessé vált kódrészletek vége