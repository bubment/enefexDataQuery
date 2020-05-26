
//GLOBÁLIS VÁLTOZÓK DEFINIÁLÁSÁNAK ELEJE
//----------------------------------------------------------------------
//Excel cellák nevei egy tömbben (A,B,C,...,ZZ)
var excelColumNames = [
    "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
    "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ",
    "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ",
    "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ",
    "EA", "EB", "EC", "ED", "EE", "EF", "EG", "EH", "EI", "EJ", "EK", "EL", "EM", "EN", "EO", "EP", "EQ", "ER", "ES", "ET", "EU", "EV", "EW", "EX", "EY", "EZ",
    "FA", "FB", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FO", "FP", "FQ", "FR", "FS", "FT", "FU", "FV", "FW", "FX", "FY", "FZ",
    "GA", "GB", "GC", "GD", "GE", "GF", "GG", "GH", "GI", "GJ", "GK", "GL", "GM", "GN", "GO", "GP", "GQ", "GR", "GS", "GT", "GU", "GV", "GW", "GX", "GY", "GZ",
    "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "HJ", "HK", "HL", "HM", "HN", "HO", "HP", "HQ", "HR", "HS", "HT", "HU", "HV", "HW", "HX", "HY", "HZ",
    "IA", "IB", "IC", "ID", "IE", "IF", "IG", "IH", "II", "IJ", "IK", "IL", "IM", "IN", "IO", "IP", "IQ", "IR", "IS", "IT", "IU", "IV", "IW", "IX", "IY", "IZ",
    "JA", "JB", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", "JM", "JN", "JO", "JP", "JQ", "JR", "JS", "JT", "JU", "JV", "JW", "JX", "JY", "JZ",
    "KA", "KB", "KC", "KD", "KE", "KF", "KG", "KH", "KI", "KJ", "KK", "KL", "KM", "KN", "KO", "KP", "KQ", "KR", "KS", "KT", "KU", "KV", "KW", "KX", "KY", "KZ",
    "LA", "LB", "LC", "LD", "LE", "LF", "LG", "LH", "LI", "LJ", "LK", "LL", "LM", "LN", "LO", "LP", "LQ", "LR", "LS", "LT", "LU", "LV", "LW", "LX", "LY", "LZ",
    "MA", "MB", "MC", "MD", "ME", "MF", "MG", "MH", "MI", "MJ", "MK", "ML", "MM", "MN", "MO", "MP", "MQ", "MR", "MS", "MT", "MU", "MV", "MW", "MX", "MY", "MZ",
    "NA", "NB", "NC", "ND", "NE", "NF", "NG", "NH", "NI", "NJ", "NK", "NL", "NM", "NN", "NO", "NP", "NQ", "NR", "NS", "NT", "NU", "NV", "NW", "NX", "NY", "NZ",
    "OA", "OB", "OC", "OD", "OE", "OF", "OG", "OH", "OI", "OJ", "OK", "OL", "OM", "ON", "OO", "OP", "OQ", "OR", "OS", "OT", "OU", "OV", "OW", "OX", "OY", "OZ",
    "PA", "PB", "PC", "PD", "PE", "PF", "PG", "PH", "PI", "PJ", "PK", "PL", "PM", "PN", "PO", "PP", "PQ", "PR", "PS", "PT", "PU", "PV", "PW", "PX", "PY", "PZ",
    "QA", "QB", "QC", "QD", "QE", "QF", "QG", "QH", "QI", "QJ", "QK", "QL", "QM", "QN", "QO", "QP", "QQ", "QR", "QS", "QT", "QU", "QV", "QW", "QX", "QY", "QZ",
    "RA", "RB", "RC", "RD", "RE", "RF", "RG", "RH", "RI", "RJ", "RK", "RL", "RM", "RN", "RO", "RP", "RQ", "RR", "RS", "RT", "RU", "RV", "RW", "RX", "RY", "RZ",
    "SA", "SB", "SC", "SD", "SE", "SF", "SG", "SH", "SI", "SJ", "SK", "SL", "SM", "SN", "SO", "SP", "SQ", "SR", "SS", "ST", "SU", "SV", "SW", "SX", "SY", "SZ",
    "TA", "TB", "TC", "TD", "TE", "TF", "TG", "TH", "TI", "TJ", "TK", "TL", "TM", "TN", "TO", "TP", "TQ", "TR", "TS", "TT", "TU", "TV", "TW", "TX", "TY", "TZ",
    "UA", "UB", "UC", "UD", "UE", "UF", "UG", "UH", "UI", "UJ", "UK", "UL", "UM", "UN", "UO", "UP", "UQ", "UR", "US", "UT", "UU", "UV", "UW", "UX", "UY", "UZ",
    "VA", "VB", "VC", "VD", "VE", "VF", "VG", "VH", "VI", "VJ", "VK", "VL", "VM", "VN", "VO", "VP", "VQ", "VR", "VS", "VT", "VU", "VV", "VW", "VX", "VY", "VZ",
    "WA", "WB", "WC", "WD", "WE", "WF", "WG", "WH", "WI", "WJ", "WK", "WL", "WM", "WN", "WO", "WP", "WQ", "WR", "WS", "WT", "WU", "WV", "WW", "WX", "WY", "WZ",
    "XA", "XB", "XC", "XD", "XE", "XF", "XG", "XH", "XI", "XJ", "XK", "XL", "XM", "XN", "XO", "XP", "XQ", "XR", "XS", "XT", "XU", "XV", "XW", "XX", "XY", "XZ",
    "YA", "YB", "YC", "YD", "YE", "YF", "YG", "YH", "YI", "YJ", "YK", "YL", "YM", "YN", "YO", "YP", "YQ", "YR", "YS", "YT", "YU", "YV", "YW", "YX", "YY", "YZ",
    "ZA", "ZB", "ZC", "ZD", "ZE", "ZF", "ZG", "ZH", "ZI", "ZJ", "ZK", "ZL", "ZM", "ZN", "ZO", "ZP", "ZQ", "ZR", "ZS", "ZT", "ZU", "ZV", "ZW", "ZX", "ZY", "ZZ",
];





// Ez a tömb tartalmaz adatokat a szerverről történő adatlekérdezésekhez
var mainItemInfo = [];

//Ez minden lekérdezéshez az URL eleje
var host = readCookie("enefexHost");
//----------------------------------------------------------------------
//GLOBÁLIS VÁLTOZÓK DEFINIÁLÁSÁNAK VÉGE



Office.onReady(function () {
    // Office is ready
    $(document).ready(function () {
        var x = document.getElementsByTagName("BODY")[0];
        x.style.display = 'block';
    });
});

Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.3')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};

//Office.onReady(function () {

//});


function onClick_logoutButton() {
    var callback = function (err, result) {
        if (err) {
            document.getElementById('logoutErrorMessage').innerHTML = err.error.message;
        }
        else {
            if (result) {
                window.location.replace("Login.html");
            }
            else {
                document.getElementById('logoutErrorMessage').innerHTML = result.error.message;

            }
        }
    };
    logoutAsync(callback);
    
}



function fogyasztasOsszesitoContainer() {


    //Globális változó a meterGroup függvény eredményének kimentéséhez.
    var meterGroupArrayResult;
    //Lekérdezésekhez szükséges URL eleje
    var host = readCookie("enefexHost");

    var errorLabel = document.getElementById('fogyasztasOsszesitoError');
    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("fogyasztas-osszesito-panel-loader", "fogyasztas-osszesito-loader", "block");

    // A függvényekben levő összes szükséges munkalapot itt kell definiálni
    var requiredSheets = ["IN_F0"];

    // Lekérdezésekhez szükséges globális paraméterek eleje
    var dateFrom = document.getElementById('kezdo_datum').value;
    var dateTo = document.getElementById('veg_datum').value;

    // Lekérdezésekhez szükséges globális paraméterek vége

    //Dátumok RegEx validációi
    if (dateRegExTest('kezdo_datum', 'veg_datum', 'fogyasztasOsszesitoError') == "RegExTestProblem") {
        setPanelLoader("fogyasztas-osszesito-panel-loader", "fogyasztas-osszesito-loader", "none");
        return;
    }

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni

    importantDisableElements = setDisableElement();
    var newDisableElements = ["kezdo_datum", "veg_datum", "fogyasztas_osszesito_meter_groups"];

    var actualDisableElements = newDisableElements.concat(importantDisableElements);

    changElementsAvailability(actualDisableElements,true);

    var meterGroup = function (callback) {

        var getMeterGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
            }
            else {
                if (result) {
                    meterGroupArrayResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                }
            }
        }

        params = {};

        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, getMeterGroupCallback);

    }

    // Munkalap létrehozó és tisztító függvény
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

    var getFogyasztasOsszesito = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A jsonDataInnerArray tömb az amivel különböző ciklusokban feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var fogyasztasOsszesitoCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("fogyasztas-osszesito-panel-loader", "fogyasztas-osszesito-loader", "none");
            }
            else {
                if (result) {

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

                    //Fejlécek betöltése a jsonDataArray-ba
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.dataTag);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(result.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            jsonDataInnerArray.push(result.data[tmpRow][requiredServerDataArray[i].dataTag]);
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    Excel.run(function (context) {

                        var sheet = context.workbook.worksheets.getItem("IN_F0");

                        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //Excel feltöltése adatokkal
                        var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        range.values = jsonDataArray;
                        range.untrack();

                        // Csak a return után lesznek láthatóak az adatok az excelben
                        boldRange.format.font.bold = true;
                        return context.sync();
                    })

                    // ---------------------EXCEL RÉSZ VÉGE --------------------

                    errorLabel.innerHTML = "";
                    errorLabel.style.display = 'none';
                    setPanelLoader("fogyasztas-osszesito-panel-loader", "fogyasztas-osszesito-loader", "none");

                    // Menü elérhetővé tétele a lekérdezés végén
                    changElementsAvailability(actualDisableElements, false);

                    //Caolan async miatt
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    setPanelLoader("fogyasztas-osszesito-panel-loader", "fogyasztas-osszesito-loader", "none");
                    changElementsAvailability(actualDisableElements, false);
                }
            }
        }

        var meterGroupList = document.getElementById('fogyasztas_osszesito_meter_groups');
        var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;
        // A szerverlekérdezéshez szükséges 'meter_gruop' paraméter értékét meghatározó változó
        var meterGroupValue;

        for (var i = 0; i < meterGroupArrayResult.length; i++) {
            if (meterGroupArrayResult[i].nev == meterGroupListSelectedText) {
                meterGroupValue = meterGroupArrayResult[i].id;
                break;
            }
        }

        var params = {};
        params["date_from"] = dateFrom;
        params["date_to"] = dateTo;
        params["meter_group"] = meterGroupValue;
        //params["date_from"] = "2019-06-01";
        //params["date_to"] = "2019-07-01";
        params["sendTo"] = "screen";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999999";

        postAsyncGetData(host + "/ebill/billing/getFogyasztasOsszesito2", params, fogyasztasOsszesitoCallback);
    }


    async.series(
        [
            meterGroup,
            workSheetHandler,
            getFogyasztasOsszesito
        ],
        function (err) {
            console.log('all finished', err);
        }
    );
}

function feldolgozottMeresekContainer() {

    var errorLabel = document.getElementById('feldolgozottMeresekError');
    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "block");

    // A függvényekben levő összes szükséges munkalapot itt kell definiálni
    var requiredSheets = ["IN_É0"];
    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");

    //Dátum RegEx validációja
    var y = document.getElementById('feldolgozottMeresekError');
    var maxRequestbeginDate = new Date();
    var currentYear = maxRequestbeginDate.getFullYear();

    if (isNaN(document.getElementById('onlyYearFilter').value) == true) {
        y.style.display = 'block';
        y.innerHTML = "A megadott év nem megfelelő formátumú. Megfelelő formátum (YYYY)";
        setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");
        return;
    }

    if (document.getElementById('onlyYearFilter').value < 2014) {
        y.style.display = 'block';
        y.innerHTML = "A megadott év 2014 előtti.";
        setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");
        return;
    }

    if (document.getElementById('onlyYearFilter').value > currentYear) {
        y.style.display = 'block';
        y.innerHTML = "A megadott év a jövőben van.";
        setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");
        return;
    }

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    importantDisableElements = setDisableElement();
    var newDisableElements = ['onlyYearFilter', 'feldolgozott_meresek_meter_groups'];
    var actualDisableElements = newDisableElements.concat(importantDisableElements);

    changElementsAvailability(actualDisableElements, true);

    // Egyéb szükséges paraméterek definiálása (input typeok tartalmai...)
    var dateMeasurementStart = document.getElementById('onlyYearFilter').value + "-01";

    //Globális változó a meterGroup függvény eredményének kimentéséhez.
    var meterGroupArrayResult;

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

    var meterGroup = function (callback) {

        var getMeterGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");
            }
            else {
                if (result) {
                    meterGroupArrayResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");
                }
            }
        }

        params = {};

        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, getMeterGroupCallback);

    }

    var getFeldolgozottMeresek = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var getFeldolgozottMeresekCallback = function (err, getFeldolgozottMeresekCallbackresult) {
            if (err) {
                errorLabel.innerHTML = err.error.message
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");
            }
            else {
                if (getFeldolgozottMeresekCallbackresult) {
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

                    //Fejlécek betöltése a jsonDataArray-ba
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.dataTag);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(getFeldolgozottMeresekCallbackresult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            jsonDataInnerArray.push(getFeldolgozottMeresekCallbackresult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    Excel.run(function (context) {

                        var sheet = context.workbook.worksheets.getItem("IN_É0");

                        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //Excel feltöltése adatokkal
                        var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        range.values = jsonDataArray;
                        range.untrack();

                        // Csak a return után lesznek láthatóak az adatok az excelben
                        boldRange.format.font.bold = true;
                        return context.sync();
                    })

                    // ---------------------EXCEL RÉSZ VÉGE --------------------

                    errorLabel.innerHTML = "";
                    errorLabel.style.display = 'none';
                    setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");

                    changElementsAvailability(actualDisableElements, false);

                    //Caolan async miatt
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A getFeldolgozottMeresekCallback resultjában lévő JSON Object hibás vagy üres";
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("feldolgozott-meresek-panel-loader", "feldolgozott-meresek-loader", "none");
                }
            }
        }

        var meterGroupList = document.getElementById('feldolgozott_meresek_meter_groups');
        var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;
        // A szerverlekérdezéshez szükséges 'meter_gruop' paraméter értékét meghatározó változó
        var meterGroupValue;

        for (var i = 0; i < meterGroupArrayResult.length; i++) {
            if (meterGroupArrayResult[i].nev == meterGroupListSelectedText) {
                meterGroupValue = meterGroupArrayResult[i].id;
                break;
            }
        }

        params = {};

        params["datum_meres_kezdete"] = dateMeasurementStart;
        params["meter_group"] = meterGroupValue;
        params["napok_mutatasa"] = "false";
        params["tankolas_is"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/summary/getFeldolgozottMeresek", params, getFeldolgozottMeresekCallback);
    }

    //HASZNOS VÁLTOZÓK

    //dataLength = Object.keys(result.data).length;
    //dataInnerLength = Object.keys(result.data[0]).length;


    async.series(
        [
            // Adatokat tartalmazó JSON lekérdezések paramétereit meghatározó függvények
            meterGroup,
            workSheetHandler,
            getFeldolgozottMeresek
            // Adatokat tartalmazó JSON lekérdezések 

        ],
        function (err) {
            console.log('allfinished', err);
        }
    )
}

function hetiJelentesKeszitoContainer() {

    var errorLabel = document.getElementById('hetiJelentesError');
    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "block");

    // A függvényekben levő összes szükséges munkalapot itt kell definiálni
    var requiredSheets = ["IN_FÖ", "IN_SzA", "IN_FG"];
    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");


    // Egyéb szükséges paraméterek definiálása (pl.: input typeok tartalmai...)

    var dateFrom = document.getElementById('heti_jelentes_kezdo_datum').value;
    var dateTo = document.getElementById('heti_jelentes_veg_datum').value;

    var dateFromHourList = document.getElementById('heti_jelentes_kezdo_ora');
    var dateFromHourSelectedText = dateFromHourList.options[dateFromHourList.selectedIndex].text;

    var dateToHourList = document.getElementById('heti_jelentes_befejezo_ora');
    var dateToHourSelectedText = dateToHourList.options[dateToHourList.selectedIndex].text;


    var savedOptionsList = document.getElementById('heti_jelentes_mentett_bealitasok');
    var savedOptionsListSelectedText
    try {
        savedOptionsListSelectedText = savedOptionsList.options[savedOptionsList.selectedIndex].text;
    } catch (e) {
        savedOptionsListSelectedText = ""
    }
    

    if (!savedOptionsListSelectedText) {
        for (var i = 0; i < requiredSheets.length; i++) {
            if (requiredSheets[i] == "IN_FG") {
                requiredSheets.splice(i, 1);
                break;
            }
        }
    }

    var csakNemMegFeleloSorokCheckBox = document.getElementById("csakNemMegFeleloSorok");

    var notShowAll;
    if (csakNemMegFeleloSorokCheckBox.checked == true) {
        notShowAll = "1";
    }
    else {
        notShowAll = "0";
    }

    var meterGroupList = document.getElementById('heti_jelentes_meter_groups');
    var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;

    var meterGroupArrayResult;
    var meterTreeArray;
    var savedOptionsArray;


    //Dátumok RegEx validációi
    if (dateRegExTest('heti_jelentes_kezdo_datum', 'heti_jelentes_veg_datum', 'hetiJelentesError') == "RegExTestProblem") {
        setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
        return;
    }

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    var newDisableElements = ['heti_jelentes_kezdo_datum', 'heti_jelentes_kezdo_ora', 'heti_jelentes_veg_datum', 'heti_jelentes_befejezo_ora', 'heti_jelentes_meter_groups', 'heti_jelentes_mentett_bealitasok', 'csakNemMegFeleloSorok'];

    importantDisableElements = setDisableElement();
    var actualDisableElements = newDisableElements.concat(importantDisableElements);

    changElementsAvailability(actualDisableElements, true);

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

    var meterGroup = function (callback) {

        var getMeterGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
            }
            else {
                if (result) {
                    meterGroupArrayResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
                }
            }
        }

        params = {};

        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, getMeterGroupCallback);

    }

    var getMeterTree = function (callback) {

        var metreGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
            }
            else {
                if (result) {
                    meterTreeArray = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
                }
            }
        }

        var params = {};

        params["node"] = "";
        params["page"] = "";

        postAsyncGetData(host + "/mdgraph/draw/getMeterTree", params, metreGroupCallback);

    }

    var getSavedGraphs = function (callback) {
        var savedGraphsCallBack = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
            }
            else {
                if (result) {
                    savedOptionsArray = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás" 
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
                }
            }
        }

        var params = {};

        params["is_public"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/mdgraph/draw/getSavedGraphs", params, savedGraphsCallBack);

    }

    var getHetiFogyasztasOsszesito = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var hetiFogyasztasOsszesitoCallback = function (err, hetiFogyasztasOsszesitoCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
            }
            else {
                if (hetiFogyasztasOsszesitoCallbackResult) {
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

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(hetiFogyasztasOsszesitoCallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            jsonDataInnerArray.push(hetiFogyasztasOsszesitoCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    Excel.run(function (context) {

                        var sheet = context.workbook.worksheets.getItem("IN_FÖ");

                        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //Excel feltöltése adatokkal
                        var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        range.values = jsonDataArray;
                        range.untrack();

                        // Csak a return után lesznek láthatóak az adatok az excelben
                        boldRange.format.font.bold = true;
                        return context.sync();
                    })

                    // ---------------------EXCEL RÉSZ VÉGE --------------------

                    errorLabel.innerHTML = "";
                    errorLabel.style.display = 'none';
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");

                    //Menü elérhetővé tétele
                    changElementsAvailability(actualDisableElements, false);

                    //Caolan async miatt
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
                }
            }
        }

        var meterGroupValue;

        for (var i = 0; i < meterGroupArrayResult.length; i++) {
            if (meterGroupArrayResult[i].nev == meterGroupListSelectedText) {
                meterGroupValue = meterGroupArrayResult[i].id;
                break;
            }
        }

        var params = {};
        params["date_from"] = dateFrom;
        params["date_to"] = dateTo;
        params["meter_group"] = meterGroupValue;
        //params["date_from"] = "2019-06-01";
        //params["date_to"] = "2019-07-01";
        params["sendTo"] = "screen";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999999";

        postAsyncGetData(host + "/ebill/billing/getFogyasztasOsszesito2", params, hetiFogyasztasOsszesitoCallback);

    }

    var getHetiAlapvonalSzamertekek = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var hetiAlapvonalSzamertekekCallback = function(err, result){
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
            }
            else {
                if (result) {
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

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(result.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            jsonDataInnerArray.push(result.data[tmpRow][requiredServerDataArray[i].dataTag]);
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    Excel.run(function (context) {

                        var sheet = context.workbook.worksheets.getItem("IN_SzA");

                        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //Excel feltöltése adatokkal
                        var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        range.values = jsonDataArray;
                        range.untrack();

                        // Csak a return után lesznek láthatóak az adatok az excelben
                        boldRange.format.font.bold = true;
                        return context.sync();
                    })

                // ---------------------EXCEL RÉSZ VÉGE --------------------

                    //Caolan async miatt
                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
                }
            }
        }

        var params = {};

        params["not_show_all"] = notShowAll;
        params["filter_date_interval"] = "1";
        params["date_from"] = dateFrom;
        params["date_to"] = dateTo;
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";
        params["sort"] = '[{ "property": "T.id", "direction": "desc" }]';

        postAsyncGetData(host + "/vstat/baseline/getTableStat", params, hetiAlapvonalSzamertekekCallback);
    }

    var getMentettBeallitasokGrafikonAdatok = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var mentettBeallitasokGrafikonAdatokCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
            }
            else {
                if (result) {
                    var extraInfoObj = result.extraInfo;
                    var extraInfoKeysArray = [];
                    for (var k in extraInfoObj) extraInfoKeysArray.push(k.replace("value", ""));

                    dataLength = Object.keys(meterTreeArray.data).length;
                    var headerArray = [];

                    extraInfoKeysArray.forEach(function (element) {

                        for (var i = 0; i < dataLength; i++) {
                            var dataSecondLevelLength = Object.keys(meterTreeArray.data[i].data).length
                            for (var j = 0; j < dataSecondLevelLength; j++) {
                                if (meterTreeArray.data[i].data[j].meter_id == element) {

                                    var elementHeaderCompatibleString = "value" + element;
                                    headerArray.push({ extraInfoKey: elementHeaderCompatibleString, extraInfoText: meterTreeArray.data[i].data[j].text });
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

                    jsonDataArray = [];
                    jsonDataInnerArray = [];

                    //Fejlécek betöltése a jsonDataArray-ba
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(result.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    var correctDateWithFormat;
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            if (requiredServerDataArray[i].dataTag == "tstamp") {
                                var d = new Date(result.data[tmpRow][requiredServerDataArray[i].dataTag]);

                                correctDateWithFormat = d.getFullYear().toString() + "-" + ((d.getMonth() + 1).toString().length == 2 ? (d.getMonth() + 1).toString() : "0" + (d.getMonth() + 1).toString()) + "-" + (d.getDate().toString().length == 2 ? d.getDate().toString() : "0" + d.getDate().toString()) + " " + (d.getHours().toString().length == 2 ? d.getHours().toString() : " " + d.getHours().toString()) + ":" + ((parseInt(d.getMinutes() / 5) * 5).toString().length == 2 ? (parseInt(d.getMinutes() / 5) * 5).toString() : "0" + (parseInt(d.getMinutes() / 5) * 5).toString()) + ":00";

                                jsonDataInnerArray.push(correctDateWithFormat);
                            }
                            else {
                                jsonDataInnerArray.push(result.data[tmpRow][requiredServerDataArray[i].dataTag]);
                            }  
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    //asd = jsonDataArray[1][0];
                    //console.log(asd);
                    //console.log(typeof asd);
                    //asd = 2;


                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    Excel.run(function (context) {

                        var sheet = context.workbook.worksheets.getItem("IN_FG");

                        var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //Excel feltöltése adatokkal
                        var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        range.values = jsonDataArray;
                        range.untrack();

                        // Csak a return után lesznek láthatóak az adatok az excelben
                        boldRange.format.font.bold = true;
                        return context.sync();
                    })

                // ---------------------EXCEL RÉSZ VÉGE --------------------

                //Caolan async miatt
                callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás";
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
                }
            }
        }

        //A datetime_from változó lesz a getGraphSeries lekérdezés datetime_from paramétere
        var datetime_from = dateFrom + ";" + dateFromHourSelectedText;

        //A datetime_to változó lesz a getGraphSeries lekérdezés datetime_to paramétere
        var datetime_to = dateTo + ";" + dateToHourSelectedText;

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

        dataLength = Object.keys(meterTreeArray.data).length

        var type_list_string = "";

        savedOptionsMetersArray.forEach(function (element) {

            for (var i = 0; i < dataLength; i++) {
                var dataSecondLevelLength = Object.keys(meterTreeArray.data[i].data).length
                for (var j = 0; j < dataSecondLevelLength; j++) {
                    if (meterTreeArray.data[i].data[j].meter_id == element) {

                        type_list_string = type_list_string.concat(meterTreeArray.data[i].data[j].data_type_id, ",");
                        i = dataLength;
                        break;
                    }
                }
            }
        });
        //A type_list_string változó lesz getGraphSeries lekérdezés type_list paramétere
        type_list_string = type_list_string.slice(0, -1);

        var params = {};

        params["datetime_from"] = datetime_from;
        params["datetime_to"] = datetime_to;
        params["meter_list"] = savedOptionsMeters;
        params["baseline_list"] = "";
        params["type_list"] = type_list_string;
        params["serie_type"] = "11";
        params["resolution"] = savedOptionsResolution;
        params["type"] = savedOptionsType;
        params["sendTo"] = "";
        params["checker"] = "0";
        params["extraInfo"] = "1";
        params["fake"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999999";

        postAsyncGetData(host + "/mdgraph/draw/getGraphSeries", params, mentettBeallitasokGrafikonAdatokCallback);
    }

    var asyncSeriesFunctionsArray = [workSheetHandler, meterGroup, getMeterTree, getSavedGraphs, getHetiAlapvonalSzamertekek, getMentettBeallitasokGrafikonAdatok, getHetiFogyasztasOsszesito];

    if (!savedOptionsListSelectedText) {
        for (var i = 0; i < asyncSeriesFunctionsArray.length; i++) {
            if (asyncSeriesFunctionsArray[i] == getMentettBeallitasokGrafikonAdatok) {
                asyncSeriesFunctionsArray.splice(i, 1);
                break;
            }
        }
    }

    async.series(
        asyncSeriesFunctionsArray,
        function (err) {
            console.log('allfinished', err);
        }
    )

}

function villamosAdminisztracioContainer() {

    var errorLabel = document.getElementById('villamosAdminisztracioError');
    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "block");
    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");
    //Ebben a tömbbe fognak kerülni az excelbe feltöltendő adatok
    var excelDataArray = [];

    var threadLimit;


    //Dátumok RegEx validációi
    //if (dateRegExTest('kezdo_datum_input', 'veg_datum_input', 'AKARMIERROR') == "RegExTestProblem") {
    //    return;
    //}

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    importantDisableElements = setDisableElement();
    var actualDisableElements = importantDisableElements
    changElementsAvailability(actualDisableElements, true);

    var csatlakozasiPontResult;
    var villanyCsoportosDijTipus;
    var katPenzeszkozok;
    var rhdDatumok;
    var rhdValues = [];
    var meddoWattosDatumok;
    var VET147Datumok;
    var HHHCSIds = [];
    var tableContentResult;
    var originalMetersResult;

    

    

    // Segéd függvények
    var getCsatlakozasiPont = function (callback) {

        var getCsatlakozasiPontCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    csatlakozasiPontResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["all"] = "1";
        params["isMasodlagos"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getHCSSzerzodes", params, getCsatlakozasiPontCallback);

    }

    var getVillanyCsoportosDijTipus = function (callback) {

        var getVillanyCsoportosDijTipusCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    villanyCsoportosDijTipus = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getVillanyCsoportosDijTipus", params, getVillanyCsoportosDijTipusCallback);
    }

    var getKatPenzeszkozValidFrom = function (callback) {

        var getKatPenzeszkozValidFromCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    katPenzeszkozok = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getKatPenzeszkozValidFrom", params, getKatPenzeszkozValidFromCallback);
    }

    var getRHDValidFrom = function (callback) {

        var getRHDValidFromCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    rhdDatumok = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getRHDValidFrom", params, getRHDValidFromCallback);
    }

    var getMeddoWattosAranyValidFrom = function (callback) {

        var getMeddoWattosAranyValidFromCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    meddoWattosDatumok = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getMeddoWattosAranyValidFrom", params, getMeddoWattosAranyValidFromCallback);
    }

    var getVET147ValidFrom = function (callback) {

        var getVET147ValidFromCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    VET147Datumok = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getVET147ValidFrom", params, getVET147ValidFromCallback);
    }

    var getTableContents = function (callback) {

        var tableContentsCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    tableContentResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["table_name"] = "ebill_villany_meresi_pont";
        params["where_to_add"] = "where_to_add";
        params["date_from"] = "date_from";
        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getTableContents", params, tableContentsCallback);
    }

    var getOriginalMeters = function (callback) {

        var originalMetersCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    originalMetersResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["unit"] = "Wh";
        params["query"] = "all";
        params["type"] = "villany";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getOriginalMeters", params, originalMetersCallback);
    }



    // Fő függvények
    var getHHSzerzodes = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];
        // Ebben az értékben tároljuk a mértékegységet tartalmazó oszlopok értékeit
        var unitColumnValue;

        var HHSzerzodesCallback = function (err, HHSzerzodesCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (HHSzerzodesCallbackResult) {
                    var requiredServerDataArray = [
                        { dataTag: "id", columnName: "A", headerText: "ID" },
                        { dataTag: "elnevezes", columnName: "B", headerText: "Elnevezés" },
                        { dataTag: "ervenyesseg_kezdete", columnName: "C", headerText: "Szerz. kezdete" },
                        { dataTag: "ervenyesseg_vege", columnName: "D", headerText: "Szerz. vége" },
                        { dataTag: "halozati_engedelyes", columnName: "E", headerText: "Hálózati engedélyes" },
                        { dataTag: "meter_identifier_watt", columnName: "F", headerText: "Mérő azonosító" },
                        { dataTag: "POD", columnName: "G", headerText: "POD" },
                        { dataTag: "consumer_tariff_type", columnName: "H", headerText: "Tarifa" },
                        { dataTag: "lekotott_teljesitmeny", columnName: "I", headerText: "Lekötött teljesítmény" },
                        { dataTag: "lekotott_teljesitmeny_mertekegyseg", columnName: "J", headerText: "Mértékegység" },
                        { dataTag: "csatlakozasi_pontok_szama", columnName: "K", headerText: "Csatlakozási Pontok száma" },
                    ];

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(HHSzerzodesCallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {

                            switch (requiredServerDataArray[i].dataTag) {
                                case "id":
                                    HHHCSIds.push(HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                                    jsonDataInnerArray.push(HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                                    break;

                                case "csatlakozasi_pontok_szama":
                                    for (var j = 0; j < csatlakozasiPontResult.data.length; j++) {
                                        if (HHSzerzodesCallbackResult.data[tmpRow].id == csatlakozasiPontResult.data[j].id) {
                                            jsonDataInnerArray.push(csatlakozasiPontResult.data[j].csatlakozasi_pontok_szama);
                                            break;
                                        }
                                    }
                                    break;

                                case "lekotott_teljesitmeny":
                                    lekotottTeljesitmeny = HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag];
                                    if (lekotottTeljesitmeny == null) {
                                        indexOfSpace = -1;
                                    }
                                    else {
                                        indexOfSpace = lekotottTeljesitmeny.indexOf(" ");
                                    }
                                    if (indexOfSpace != -1) {
                                        jsonDataInnerArray.push(lekotottTeljesitmeny.substr(0, indexOfSpace))
                                        unitColumnValue = lekotottTeljesitmeny.substr(indexOfSpace + 1, lekotottTeljesitmeny.length)
                                    }
                                    else {
                                        jsonDataInnerArray.push("undefined");
                                        unitColumnValue = "undefined";
                                    }
                                    break;

                                case "lekotott_teljesitmeny_mertekegyseg":
                                    jsonDataInnerArray.push(unitColumnValue)
                                    unitColumnValue = "";
                                    break;

                                default:
                                    jsonDataInnerArray.push(HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                            }
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }


                    excelDataArray.push(
                            {
                                "sheetName": "HHHCS szerződések",
                                "data": jsonDataArray,
                            }
                        )
                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    //Excel.run(function (context) {

                    //    var sheet = context.workbook.worksheets.getItem("HHHCS szerződések");

                    //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                    //    //Excel feltöltése adatokkal
                    //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    //    range.values = jsonDataArray;
                    //    range.untrack();

                    //    // Csak a return után lesznek láthatóak az adatok az excelben
                    //    boldRange.format.font.bold = true;
                    //    return context.sync();
                    //})

                    // ---------------------EXCEL RÉSZ VÉGE --------------------
                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        var params = {};
        params["all"] = "1";
        params["isMasodlagos"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getHHSzerzodes", params, HHSzerzodesCallback);

    }

    var getOperativTeljesitmeny = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var params = {};
        params["isMasodlagos"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "id", columnName: "A", headerText: "ID" }, // nem ennek a lekérdezésnek az ID-ja, hanem a hozzátartozó szerződésé
            { dataTag: "value", columnName: "B", headerText: "Érték" },
            { dataTag: "value_mertekegyseg", columnName: "C", headerText: "Mértékegység" },
            { dataTag: "ervenyesseg_kezdete", columnName: "D", headerText: "Érvényesség kezdete" },
            { dataTag: "ervenyesseg_vege", columnName: "E", headerText: "Érvényesség vége" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var operativTeljesitmenyek = function (item, innerCallback) {

            var operativTeljesitmenyekCallback = function (err, operativTeljesitmenyekCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = operativTeljesitmenyekCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                switch (requiredServerDataArray[j].dataTag) {
                                    case "id":
                                        jsonDataInnerArray.push(item);
                                        break;
                                    case "value":
                                        ertek = operativTeljesitmenyekCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (ertek == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = ertek.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(ertek.substr(0, indexOfSpace))
                                            unitColumnValue = ertek.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "value_mertekegyseg":
                                        jsonDataInnerArray.push(unitColumnValue)
                                        unitColumnValue = "";
                                        break;

                                    default:
                                        jsonDataInnerArray.push(operativTeljesitmenyekCallbackResult[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                            
                        }
                    }
                    requestCounter++;

                    //if (item == HHHCSIds[HHHCSIds.length - 1]) {
                    if (requestCounter == HHHCSIds.length) {
                        dataLength = jsonDataArray.length - 1;

                        excelDataArray.push(
                            {
                                "sheetName": "Operatív teljesítmény",
                                "data": jsonDataArray,
                            }
                        )

                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("Operatív teljesítmény");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["operativ_szerzodes_id"] = item;
            postAsyncGetData(host + "/ebill/contract/Get_ebill_operativ_szerzodes", params, operativTeljesitmenyekCallback);
        };

        async.eachLimit(
            HHHCSIds,
            threadLimit,
            operativTeljesitmenyek,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getVillanyCsoportosDij = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var villanyCsoportosDijArray = [];
        for (var i = 0; i < villanyCsoportosDijTipus.data.length; i++) {
            villanyCsoportosDijArray.push(villanyCsoportosDijTipus.data[i].id);
        }

        if (villanyCsoportosDijArray.length == 0) {
            villanyCsoportosDijArray.push("");
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "id", columnName: "A", headerText: "ID" },
            { dataTag: "tipus", columnName: "B", headerText: "Csoport" },
            { dataTag: "ervenyesseg_kezdete", columnName: "C", headerText: "Érvényesség kezdete" },
            { dataTag: "ervenyesseg_vege", columnName: "D", headerText: "Érvényesség vége" },
            { dataTag: "energia_ar_mod_kwh", columnName: "E", headerText: "Érték" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var villanyCsoportosDij = function (item, innerCallback) {

            var villanyCsoportosDijCallback = function (err, villanyCsoportosDijCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = villanyCsoportosDijCallbackResult.data.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                        jsonDataInnerArray.push(villanyCsoportosDijCallbackResult.data[i][requiredServerDataArray[j].dataTag]);
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                            
                        }
                    }
                    requestCounter++;

                    //if (item == villanyCsoportosDijArray[villanyCsoportosDijArray.length - 1]) {
                    if (requestCounter == villanyCsoportosDijArray.length) {
                        dataLength = jsonDataArray.length - 1;

                        excelDataArray.push(
                            {
                                "sheetName": "Csoportos díj módosító",
                                "data": jsonDataArray,
                            }
                        )

                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("Csoportos díj módosító");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["tipus"] = item;
            postAsyncGetData(host + "/ebill/contract/getVillanyCsoportosDij", params, villanyCsoportosDijCallback);
        };

        async.eachLimit(
            villanyCsoportosDijArray,
            threadLimit,
            villanyCsoportosDij,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getKatPenzeszkozValues = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var katPenzeszkozokArray = [];
        for (var i = 0; i < katPenzeszkozok.length; i++) {
            katPenzeszkozokArray.push(katPenzeszkozok[i].ervenyesseg_kezdete);
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "ervenyesseg_kezdete", columnName: "A", headerText: "Érvényesség kezdete" },
            { dataTag: "ervenyesseg_vege", columnName: "B", headerText: "Érvényesség vége" },
            { dataTag: "kat_penzeszkoz_egysegar", columnName: "C", headerText: "KÁT pénzeszköz egységár (Ft/kWh)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var katPenzeszkozValues = function (item, innerCallback) {

            var katPenzeszkozValuesCallback = function (err, katPenzeszkozValuesCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = katPenzeszkozValuesCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {

                                switch (requiredServerDataArray[j].dataTag) {
                                    case "kat_penzeszkoz_egysegar":
                                        penzeszkozEgysegar = katPenzeszkozValuesCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (penzeszkozEgysegar == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = penzeszkozEgysegar.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(penzeszkozEgysegar.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    default:
                                        jsonDataInnerArray.push(katPenzeszkozValuesCallbackResult[i][requiredServerDataArray[j].dataTag]);
                                }
                                
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }
                    requestCounter++;
                    if (requestCounter == katPenzeszkozokArray.length) {
                    //if (item == katPenzeszkozokArray[katPenzeszkozokArray.length - 1]) {
                        dataLength = jsonDataArray.length - 1;


                        excelDataArray.push(
                            {
                                "sheetName": "KÁT pénzeszköz",
                                "data": jsonDataArray,
                            }
                        )
                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("KÁT pénzeszköz");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item;
            postAsyncGetData(host + "/ebill/admin/getKatPenzeszkozMertekegyseggelValues", params, katPenzeszkozValuesCallback);
        };

        async.eachLimit(
            katPenzeszkozokArray,
            threadLimit,
            katPenzeszkozValues,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getAllandoRendszerhasznalatiDijak = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var rhdDatumokArray = [];
        for (var i = 0; i < rhdDatumok.length; i++) {
            rhdDatumokArray.push(rhdDatumok[i].ervenyesseg_kezdete);
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "ervenyesseg_kezdete", columnName: "A", headerText: "Érvényesség kezdete" },
            { dataTag: "ervenyesseg_vege", columnName: "B", headerText: "Érvényesség vége" },
            { dataTag: "atviteli_rendszeriranyitasi_dij", columnName: "C", headerText: "Átviteli rendszerirányítási díj (Ft/kWh)" },
            { dataTag: "rendszerszintu_szolgaltatasi_dij", columnName: "D", headerText: "Rendszerszintű szolgáltatási díj (Ft/kWh)" },
            { dataTag: "kozvilagitasi_elosztasi_dij", columnName: "E", headerText: "Közvilágítási elosztási díj (Ft/kWh)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var AllandoRendszerhasznalatiDijak = function (item, innerCallback) {

            var AllandoRendszerhasznalatiDijakCallback = function (err, AllandoRendszerhasznalatiDijakCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
                else {
                    //RHD adatok kimetése, hogy a következő fő függvény is tudja használni
                    // EZ IS FUNKCIONÁL MINT SEGÉDFÜGGVÉNY!!!!
                    rhdValues.push(AllandoRendszerhasznalatiDijakCallbackResult);
                    
                    dataLength = AllandoRendszerhasznalatiDijakCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {

                                switch (requiredServerDataArray[j].dataTag) {
                                    case "atviteli_rendszeriranyitasi_dij":
                                        value = AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "kozvilagitasi_elosztasi_dij":
                                        value = AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "rendszerszintu_szolgaltatasi_dij":
                                        value = AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    default:
                                        jsonDataInnerArray.push(AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray);  
                        }
                    }
                    requestCounter++;

                    //if (item == rhdDatumok[rhdDatumok.length - 1].ervenyesseg_kezdete) {
                    if (requestCounter == rhdDatumok.length) {
                        dataLength = jsonDataArray.length - 1;

                        excelDataArray.push(
                            {
                                "sheetName": "RHD azonos",
                                "data": jsonDataArray,
                            }
                        )


                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("RHD azonos");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item;
            postAsyncGetData(host + "/ebill/admin/getRHDAllandoValues_mertekegyseggel", params, AllandoRendszerhasznalatiDijakCallback);
        };

        async.eachLimit(
            rhdDatumokArray,
            threadLimit,
            AllandoRendszerhasznalatiDijak,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getRendszerhasznalatiDijak = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var rhdValuesArray = [];
        for (var i = 0; i < rhdValues.length; i++) {
            rhdValuesArray.push({ "kedo_datum": rhdValues[i][0].ervenyesseg_kezdete, "befejezo_datum": rhdValues[i][0].ervenyesseg_vege });
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "ervenyesseg_kezdete", columnName: "A", headerText: "Érvényesség kezdete" }, // Ennek az értéke nem a lekérdezésből jön
            { dataTag: "ervenyesseg_vege", columnName: "B", headerText: "Érvényesség vége" }, // Ennek az értéke nem a lekérdezésből jön
            { dataTag: "fogyaszto_tipus", columnName: "C", headerText: "Tarifa típus" },
            { dataTag: "elosztoi_alapdij", columnName: "D", headerText: "Elosztói alapdíj (Ft/csatl.p/év)" },
            { dataTag: "elosztoi_teljesitmeny_dij", columnName: "E", headerText: "Elosztói teljesítménydí (Ft/kW/év)" },
            { dataTag: "elosztoi_forgalmi_dij", columnName: "F", headerText: "Elosztói forgalmi díj (Ft/kWh)" },
            { dataTag: "elosztoi_meddo_energia_dij", columnName: "G", headerText: "Elosztói meddő energia díj (Ft/kVArh)" },
            { dataTag: "elosztoi_veszteseg_dij", columnName: "H", headerText: "Elosztói veszteség díj (Ft/kWh)" },
            { dataTag: "elosztoi_menetrend_kiegyensulyozasi_dij", columnName: "I", headerText: "Menetrend kiegyens. díj (Ft/kWh)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var rendszerHasznalatiDijak = function (item, innerCallback) {

            var rendszerHasznalatiDijakCallBack = function (err, rendszerHasznalatiDijakCallBackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = rendszerHasznalatiDijakCallBackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {

                                switch (requiredServerDataArray[j].dataTag) {
                                    case "ervenyesseg_kezdete":
                                        jsonDataInnerArray.push(item.kedo_datum);
                                        break;
                                    case "ervenyesseg_vege":
                                        jsonDataInnerArray.push(item.befejezo_datum);
                                        break;
                                    case "elosztoi_alapdij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_forgalmi_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_meddo_energia_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_menetrend_kiegyensulyozasi_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_teljesitmeny_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_veszteseg_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;

                                    default:
                                        jsonDataInnerArray.push(rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray); 
                        }
                    }
                    requestCounter++;

                    //if (item.kedo_datum == rhdValuesArray[rhdValuesArray.length - 1].kedo_datum) {
                    if (requestCounter == rhdValuesArray.length) {
                        dataLength = jsonDataArray.length - 1;


                        excelDataArray.push(
                            {
                                "sheetName": "RHD tarifafüggő",
                                "data": jsonDataArray,
                            }
                        )
                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("RHD tarifafüggő");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item.kedo_datum;
            postAsyncGetData(host + "/ebill/admin/getRHDValues_mertekegyseggel", params, rendszerHasznalatiDijakCallBack);
        };

        async.eachLimit(
            rhdValuesArray,
            threadLimit,
            rendszerHasznalatiDijak,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getMeddoWattosValues = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;


        var meddoWattosArray = [];
        for (var i = 0; i < meddoWattosDatumok.length; i++) {
            meddoWattosArray.push(meddoWattosDatumok[i].ervenyesseg_kezdete);
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "tarifatipus", columnName: "A", headerText: "Tarifa típus" },
            { dataTag: "ervenyesseg_kezdete", columnName: "B", headerText: "Érvényesség kezdete" },
            { dataTag: "arany", columnName: "C", headerText: "Arány (%)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var meddoWattosValues = function (item, innerCallback) {

            var meddoWattosValuesCallback = function (err, meddoWattosValuesCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = meddoWattosValuesCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                        jsonDataInnerArray.push(meddoWattosValuesCallbackResult[i][requiredServerDataArray[j].dataTag]);
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }
                    requestCounter++;
                    //if (item == meddoWattosArray[meddoWattosArray.length - 1]) {
                    if (requestCounter == meddoWattosArray.length) {
                        dataLength = jsonDataArray.length - 1;

                        excelDataArray.push(
                            {
                                "sheetName": "Meddő-Wattos arány",
                                "data": jsonDataArray,
                            }
                        )

                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("Meddő-Wattos arány");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item;
            postAsyncGetData(host + "/ebill/admin/getMeddoWattosValues", params, meddoWattosValuesCallback);
        };

        async.eachLimit(
            meddoWattosArray,
            threadLimit,
            meddoWattosValues,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getVET147Values = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var VET147Array = [];
        for (var i = 0; i < VET147Datumok.length; i++) {
            VET147Array.push(VET147Datumok[i].ervenyesseg_kezdete);
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "ervenyesseg_kezdete", columnName: "A", headerText: "Érvényesség kezdete" },
            { dataTag: "ervenyesseg_vege", columnName: "B", headerText: "Érvényesség vége" },
            { dataTag: "kedvezmenyes_tamogatas_dij", columnName: "C", headerText: "Kedvezményes áru VE. támogatás díja (Ft/kWh)" },
            { dataTag: "szenipari_szerkezetatalakitasi_dij", columnName: "D", headerText: "Szénipari szerkezetátalakítási díj (Ft/kWh)" },
            { dataTag: "kapcsolt_termelesszerk_atalakitasi_dij", columnName: "E", headerText: "Kapcsolt termelésszerk.. átalakítási díj (Ft/kWh)" },
            { dataTag: "energia_ado", columnName: "F", headerText: "Energia adó (Ft/kWh)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var VET147Values = function (item, innerCallback) {

            var VET147ValuesCallback = function (err, VET147ValuesCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = VET147ValuesCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                        jsonDataInnerArray.push(VET147ValuesCallbackResult[i][requiredServerDataArray[j].dataTag]);
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }
                    requestCounter++;
                    
                    //if (item == VET147Array[VET147Array.length - 1]) {
                    if (requestCounter == VET147Array.length) {
                        dataLength = jsonDataArray.length - 1;


                        excelDataArray.push(
                            {
                                "sheetName": "VET147",
                                "data": jsonDataArray,
                            }
                        )

                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("VET147");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item;
            postAsyncGetData(host + "/ebill/admin/getVET147Values", params, VET147ValuesCallback);
        };

        async.eachLimit(
            VET147Array,
            threadLimit,
            VET147Values,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getKereskedelmiSzerzodes = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var kereskedekmiSzerzodesCallback = function (err, kereskedekmiSzerzodesCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (kereskedekmiSzerzodesCallbackResult) {
                    var requiredServerDataArray = [
                        { dataTag: "id", columnName: "A", headerText: "ID" },
                        { dataTag: "elnevezes", columnName: "B", headerText: "Elnevezés" },
                        { dataTag: "ervenyesseg_kezdete", columnName: "C", headerText: "Szerz. kezdete" },
                        { dataTag: "ervenyesseg_vege", columnName: "D", headerText: "Szerz. vége" },
                        { dataTag: "villamosenergia_kereskedo_hosszu_nev", columnName: "E", headerText: "Kereskedelmi  engedélyes" },
                        { dataTag: "meter_identifier_watt", columnName: "F", headerText: "Mérő azonosító" },
                        { dataTag: "POD", columnName: "G", headerText: "POD" },
                        { dataTag: "villamos_energia_ar_kwh", columnName: "H", headerText: "Áramdíj" },
                        { dataTag: "csoportos_dij_tipus_id", columnName: "I", headerText: "Díjmódosító csoport" },
                    ];

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(kereskedekmiSzerzodesCallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            switch (requiredServerDataArray[i].dataTag) {

                                case "villamos_energia_ar_kwh":
                                    villamosEnergiaAr = kereskedekmiSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag];
                                    if (villamosEnergiaAr == null) {
                                        indexOfSpace = -1;
                                    }
                                    else {
                                        indexOfSpace = villamosEnergiaAr.indexOf(" ");
                                    }
                                    if (indexOfSpace != -1) {
                                        jsonDataInnerArray.push(villamosEnergiaAr.substr(0, indexOfSpace))
                                        //unitColumnValue = villamosEnergiaAr.substr(indexOfSpace + 1, villamosEnergiaAr.length)
                                    }
                                    else {
                                        jsonDataInnerArray.push("undefined");
                                        //unitColumnValue = "undefined";
                                    }
                                    break;

                                default:
                                    jsonDataInnerArray.push(kereskedekmiSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);

                            }
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    excelDataArray.push(
                        {
                            "sheetName": "Kereskedelmi szerződések",
                            "data": jsonDataArray,
                        }
                    )

                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    //Excel.run(function (context) {

                    //    var sheet = context.workbook.worksheets.getItem("Kereskedelmi szerződések");

                    //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                    //    //Excel feltöltése adatokkal
                    //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    //    range.values = jsonDataArray;
                    //    range.untrack();

                    //    // Csak a return után lesznek láthatóak az adatok az excelben
                    //    boldRange.format.font.bold = true;
                    //    return context.sync();
                    //})

                    // ---------------------EXCEL RÉSZ VÉGE --------------------

                    //errorLabel.innerHTML = "";
                    //errorLabel.style.display = 'none';

                    ////Menü elérhetővé tétele
                    //changElementsAvailability(actualDisableElements, false);

                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }


        var params = {};
        params["all"] = "1";
        params["isMasodlagos"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";




        postAsyncGetData(host + "/ebill/contract/getKereskedelmiSzerzodes", params, kereskedekmiSzerzodesCallback);

    }

    var getVillanySzerzodesVet147 = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];
        // Ebben az értékben tároljuk a mértékegységet tartalmazó oszlopok értékeit
        var unitColumnValue;

        var villanySzerzodesVet147Callback = function (err, villanySzerzodesVet147CallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
            }
            else {
                if (villanySzerzodesVet147CallbackResult) {
                    var requiredServerDataArray = [
                        { dataTag: "id", columnName: "A", headerText: "ID" },
                        { dataTag: "elnevezes", columnName: "B", headerText: "Elnevezés" },
                        { dataTag: "ervenyesseg_kezdete", columnName: "C", headerText: "Érvényesség kezdete" },
                        { dataTag: "ervenyesseg_vege", columnName: "D", headerText: "Érvényesség vége" },
                        { dataTag: "tovabbado_elszamolasi_meroje_id", columnName: "E", headerText: "Továbbadó elszámolási mérője" },//tableContentResult-ból jön
                        { dataTag: "tovabbszamlazott_mero_id", columnName: "F", headerText: "Továbbszámlázott mérő" }, // originalMetersResult-ból jön
                        { dataTag: "alapdij_felosztas", columnName: "G", headerText: "Alapdíj felosztás típusa" }, // 5 = Fajlagos, 4 = Fix költség, 1 = Százalékos, default = HIBA 
                        { dataTag: "alapdij_fix_koltseg_ertek", columnName: "H", headerText: "Alapdíj fix költség érték" },
                        { dataTag: "alapdij_szazalekos_ertek", columnName: "I", headerText: "Alapdíj százalékos érték" },
                        { dataTag: "almero_nelkuli", columnName: "J", headerText: "Almérő nélküli" },
                        { dataTag: "energia_felosztas", columnName: "K", headerText: "Energia felosztás típusa" }, //1= Százalékos, 2 = Almérő alapján (főmérőhöz), 3 = Almérő alapján (almérők összege), 4 = Fix költség
                        { dataTag: "energia_fix_koltseg_ertek", columnName: "L", headerText: "Energia fix költség érték" },
                        { dataTag: "energia_szazalekos_ertek", columnName: "M", headerText: "Energia százalékos érték" },
                        { dataTag: "maganvezetek_hasznalati_dij", columnName: "N", headerText: "Magánvezeték használati díj" },
                        { dataTag: "megjegyzes", columnName: "O", headerText: "Megjegyzés" },
                    ];

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(villanySzerzodesVet147CallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {

                            switch (requiredServerDataArray[i].dataTag) {
                                case "tovabbado_elszamolasi_meroje_id":
                                    for (var j = 0; j < tableContentResult.length; j++) {
                                        if (tableContentResult[j].id == villanySzerzodesVet147CallbackResult.data[tmpRow].tovabbado_elszamolasi_meroje_id) {
                                            reparableValue = tableContentResult[j].watt_meter_nev;
                                            if (typeof reparableValue == "string") {
                                                subStringStart = reparableValue.indexOf("-") + 2;
                                                subStringLength = reparableValue.indexOf(" ", subStringStart);
                                                reparableValue = reparableValue.substring(subStringStart, subStringLength);
                                            }
                                            else {
                                                reparableValue = "undefined"
                                            }
                                            jsonDataInnerArray.push(reparableValue);
                                            break;
                                        }
                                    }
                                    break;
                                case "tovabbszamlazott_mero_id":
                                    for (var k = 0; k < originalMetersResult.length; k++) {
                                        if (originalMetersResult[k].id == villanySzerzodesVet147CallbackResult.data[tmpRow].tovabbszamlazott_mero_id) {
                                            value = originalMetersResult[k].identifier;
                                            jsonDataInnerArray.push(value);
                                            break;
                                        }
                                    }
                                    break;
                                case "alapdij_felosztas":
                                    alapDijFeloszttas = villanySzerzodesVet147CallbackResult.data[tmpRow].alapdij_felosztas;
                                    switch (alapDijFeloszttas) {
                                        case "1":
                                            value = "Százalékos";
                                            break;
                                        case "4":
                                            value = "Fix költség";
                                            break;
                                        case "5":
                                            value = "Fajlagos";
                                            break;
                                        default:
                                            value = "undefined";
                                    }
                                    jsonDataInnerArray.push(value);
                                    break;
                                case "energia_felosztas":
                                    energiaFelosztas = villanySzerzodesVet147CallbackResult.data[tmpRow].energia_felosztas;
                                    switch (energiaFelosztas) {
                                        case "1":
                                            value = "Százalékos";
                                            break;
                                        case "2":
                                            value = "Almérő alapján (főmérőhöz)";
                                            break;
                                        case "3":
                                            value = "Almérő alapján (almérők összege)";
                                            break;
                                        case "4":
                                            value = "Fix költség";
                                            break;
                                        default:
                                            value = "undefined";
                                    }
                                    jsonDataInnerArray.push(value);
                                    break;


                                default:
                                    jsonDataInnerArray.push(villanySzerzodesVet147CallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                            }
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    excelDataArray.push(
                        {
                            "sheetName": "Villanyszerződés Vet147",
                            "data": jsonDataArray,
                        }
                    )

                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    //Excel.run(function (context) {

                    //    var sheet = context.workbook.worksheets.getItem("Villanyszerződés Vet147");

                    //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                    //    //Excel feltöltése adatokkal
                    //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    //    range.values = jsonDataArray;
                    //    range.untrack();

                    //    // Csak a return után lesznek láthatóak az adatok az excelben
                    //    boldRange.format.font.bold = true;
                    //    return context.sync();
                    //})

                    // ---------------------EXCEL RÉSZ VÉGE --------------------
                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                }
            }
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getVillanySzerzodesVet147", params, villanySzerzodesVet147Callback);

    }

    //Függvény ami kezeli az excelt
    //Munkalapokat hoz létre, munkalapokat tisztít és feétölti a megfelelő munkalapokat adatokkal
    var workSheetHandler = function (callback_lvl1) {

        var clearableSheet = [];
        var addableSheet = [];

        var separateWorksheets = function (callback_lvl2) {
            // Ez a függvény a lekérdezéshez szükséges munkalapokat két külön tömbe teszi.
            // A clearableSheet tömbbe teszi a már létező munkalapok nevét
            // Az addableSheet tömbbe teszi a létrehozandó munkalapok nevét
            Excel.run(function (context) {
                var worksheets = context.workbook.worksheets;
                worksheets.load('name');
                return context.sync()
                    .then(function () {
                        var sheetFound;
                        for (var i = 0; i < excelDataArray.length; i++) {
                            sheetFound = false;
                            for (var j = 0; j < worksheets.items.length; j++) {
                                if (excelDataArray[i].sheetName == worksheets.items[j].name) {
                                    sheetFound = true;
                                    clearableSheet.push(worksheets.items[j].name);
                                    break;
                                }
                            }
                            if (sheetFound) {
                                continue;
                            }
                            else {
                                addableSheet.push(excelDataArray[i].sheetName);
                            }

                        }
                        callback_lvl2();
                    });
            })

        }

        var clearSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var addSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var loadDataToSheets = function (callback_lvl2) {
            if (excelDataArray) {
                Excel.run(function (context) {
                    var sheet;
                    var range;
                    var columnName;
                    var rowValue;

                    ////Munkalap nevének meghatározása
                    //sheet = context.workbook.worksheets.getItem("HHHCS szerződések");
                    ////Adatokkal feltöltendő tartomány meghatározésa
                    //range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    ////Adatok feltöltése
                    //range.values = jsonDataArray;
                    //range.untrack();

                    for (var i = 0; i < excelDataArray.length; i++) {
                        sheet = context.workbook.worksheets.getItem(excelDataArray[i].sheetName);
                        columnName = excelColumNames[excelDataArray[i].data[0].length - 1];
                        rowValue = excelDataArray[i].data.length;

                        range = sheet.getRange("A1:" + columnName + rowValue);
                        range.values = excelDataArray[i].data;
                        range.untrack();

                    }


                    return context.sync()
                        .then(function () {
                            errorLabel.innerHTML = "";
                            errorLabel.style.display = "none";
                            changElementsAvailability(actualDisableElements, false);
                            setPanelLoader("villamos-adminisztracio-panel-loader", "villamos-adminisztracio-loader", "none");
                            callback_lvl2();
                            callback_lvl1();
                        });
                });
            }
        }

        async.series(
            [
                separateWorksheets,
                clearSheets,
                addSheets,
                loadDataToSheets
            ],
            function (err) {
                console.log('all finished', err);
            }
        );
    }
    


    async.series(
        [
            
            getCsatlakozasiPont,
            getVillanyCsoportosDijTipus,
            getKatPenzeszkozValidFrom,
            getRHDValidFrom,
            getMeddoWattosAranyValidFrom,
            getVET147ValidFrom,
            getTableContents,
            getOriginalMeters,
            getHHSzerzodes,
            getOperativTeljesitmeny,
            getVillanyCsoportosDij,
            getKatPenzeszkozValues,
            getAllandoRendszerhasznalatiDijak,
            getRendszerhasznalatiDijak,
            getMeddoWattosValues,
            getVET147Values,
            getVillanySzerzodesVet147,
            //Az utolso maradjon utolso
            getKereskedelmiSzerzodes,
            workSheetHandler,
        ],
        function (err) {
            console.log('allfinished', err);
        }
    )

}

function szamlaOsszesitoContainer() {

    var errorLabel = document.getElementById('szamlaOsszesitoError');
    var maxRequestbeginDate = new Date();
    var currentYear = maxRequestbeginDate.getFullYear();

    // onlyYearFilter helyett más id
    if (isNaN(document.getElementById('szamlaOsszesitoYearFilter').value) == true) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A megadott év nem megfelelő formátumú. Megfelelő formátum (YYYY)"
        return;
    }

    if (document.getElementById('szamlaOsszesitoYearFilter').value > currentYear) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A megadott év a jövőben van."
        return;
    }


    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("szamla-osszesito-panel-loader", "szamla-osszesito-loader", "block");

    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");
    //Ebben a tömbbe fognak kerülni az excelbe feltöltendő adatok
    var excelDataArray = []

    //var threadLimit;

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    importantDisableElements = setDisableElement();
    var newDisableElements = ["szamlaOsszesitoYearFilter"];
    var actualDisableElements = newDisableElements.concat(importantDisableElements);
    changElementsAvailability(actualDisableElements, true);


    // Fő függvények

    var getSzamlaOsszesito = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];


        var SzamlaOsszesitoCallback = function (err, SzamlaOsszesitoCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("szamla-osszesito-panel-loader", "szamla-osszesito-loader", "none");
            }
            else {
                if (SzamlaOsszesitoCallbackResult) {
                    var requiredServerDataArray = [
                        { dataTag: "statusz", columnName: "A", headerText: "Státusz" },
                        { dataTag: "egyeb_azonosito_1", columnName: "B", headerText: "Beérkezett" },
                        { dataTag: "egyeb_azonosito_2", columnName: "C", headerText: "Belső azonosító" },
                        { dataTag: "epulet_azonosito", columnName: "D", headerText: "Épület azonositó" },
                        { dataTag: "meter_name", columnName: "E", headerText: "Mérés neve" },
                        { dataTag: "meres_tipus", columnName: "F", headerText: "Mérés típus" },
                        { dataTag: "szamlatipus_nev", columnName: "G", headerText: "Számla típus" },
                        { dataTag: "szolgaltatoi_szamlaszam", columnName: "H", headerText: "Szolgáltatói számlaszám" },
                        { dataTag: "idoszak_kezdete", columnName: "I", headerText: "Leolvasás kezdete" },
                        { dataTag: "idoszak_vege", columnName: "J", headerText: "Leolvasás vége" },
                        { dataTag: "hatasos_fogyasztas_fogyasztas", columnName: "K", headerText: "Fogyasztás" },
                        { dataTag: "hatasos_fogyasztas_fogyasztas_mertekegyseg", columnName: "L", headerText: "[]" },
                        { dataTag: "AHK", columnName: "M", headerText: "AHK" },
                        { dataTag: "netto_osszeg_mertekegyseg", columnName: "N", headerText: "[]" },
                        { dataTag: "szamla_netto_osszeg", columnName: "O", headerText: "Nettó számla" },
                        { dataTag: "netto_osszeg_mertekegyseg", columnName: "P", headerText: "[]" },
                        { dataTag: "szamla_afa", columnName: "Q", headerText: "ÁFA" },
                        { dataTag: "netto_osszeg_mertekegyseg", columnName: "R", headerText: "[]" },
                        { dataTag: "szamla_brutto_osszeg", columnName: "S", headerText: "Bruttó számla" },
                        { dataTag: "netto_osszeg_mertekegyseg", columnName: "T", headerText: "[]" },
                        { dataTag: "elhatarolas_brutto_osszeg", columnName: "U", headerText: "Elhatárolás" }
                    ];

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(SzamlaOsszesitoCallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    var actDateStr;
                    var actYearInt;
                    var filterYear = document.getElementById('szamlaOsszesitoYearFilter').value;
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        actDateStr = SzamlaOsszesitoCallbackResult.data[tmpRow].idoszak_kezdete;
                        actYearInt = parseInt(actDateStr.substring(0, 4));
                        //(t-1) és (t) az aktuális panelról jöjjön
                        if (actYearInt == filterYear || actYearInt == (filterYear-1)) {
                            jsonDataInnerArray = [];
                            for (var i = 0; i < dataInnerLength; i++) {
                                jsonDataInnerArray.push(SzamlaOsszesitoCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }


                    }


                    excelDataArray.push(
                        {
                            "sheetName": "Számlaösszesítő",
                            "data": jsonDataArray,
                        }
                    )
                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    //Excel.run(function (context) {

                    //    var sheet = context.workbook.worksheets.getItem("HHHCS szerződések");

                    //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                    //    //Excel feltöltése adatokkal
                    //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    //    range.values = jsonDataArray;
                    //    range.untrack();

                    //    // Csak a return után lesznek láthatóak az adatok az excelben
                    //    boldRange.format.font.bold = true;
                    //    return context.sync();
                    //})

                    // ---------------------EXCEL RÉSZ VÉGE --------------------
                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("szamla-osszesito-panel-loader", "szamla-osszesito-loader", "none");
                }
            }
        }

        var params = {};
        params["table_name"] = "ebill_villany_KAT";
        params["where_to_add"] = "where_to_add";
        params["szamla_type"] = "ebill_foldgaz_szamla_kereskedelmi";
        params["service_type_union"] = "true";
        params["date_from"] = "date_from_KAT_FO";
        //params["date_from"] = "2018-01-01";
        params["page"] = "1";
        params["start"] = "0";
        params["elhatarolas"] = "1";
        params["limit"] = "99999";
        //filter: [{ "property": "idoszak_kezdete", "value": "2018-01-01" }]
        //params["filter"] = [{ "property": "idoszak_kezdete", "value": "2018-01-01" }];

        postAsyncGetData(host + "/ebill/billing/getSzamlaOsszesitoSzamla", params, SzamlaOsszesitoCallback);

    }

    //Függvény ami kezeli az excelt

    var workSheetHandler = function (callback_lvl1) {

        var clearableSheet = [];
        var addableSheet = [];

        var separateWorksheets = function (callback_lvl2) {
            // Ez a függvény a lekérdezéshez szükséges munkalapokat két külön tömbe teszi.
            // A clearableSheet tömbbe teszi a már létező munkalapok nevét
            // Az addableSheet tömbbe teszi a létrehozandó munkalapok nevét
            Excel.run(function (context) {
                var worksheets = context.workbook.worksheets;
                worksheets.load('name');
                return context.sync()
                    .then(function () {
                        var sheetFound;
                        for (var i = 0; i < excelDataArray.length; i++) {
                            sheetFound = false;
                            for (var j = 0; j < worksheets.items.length; j++) {
                                if (excelDataArray[i].sheetName == worksheets.items[j].name) {
                                    sheetFound = true;
                                    clearableSheet.push(worksheets.items[j].name);
                                    break;
                                }
                            }
                            if (sheetFound) {
                                continue;
                            }
                            else {
                                addableSheet.push(excelDataArray[i].sheetName);
                            }

                        }
                        callback_lvl2();
                    });
            })

        }

        var clearSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var addSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var loadDataToSheets = function (callback_lvl2) {
            if (excelDataArray) {
                Excel.run(function (context) {

                    context.application.suspendApiCalculationUntilNextSync();

                    var sheet;
                    var range;
                    var columnName;
                    var rowValue;

                    ////Munkalap nevének meghatározása
                    //sheet = context.workbook.worksheets.getItem("HHHCS szerződések");
                    ////Adatokkal feltöltendő tartomány meghatározésa
                    //range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    ////Adatok feltöltése
                    //range.values = jsonDataArray;
                    //range.untrack();

                    for (var i = 0; i < excelDataArray.length; i++) {
                        sheet = context.workbook.worksheets.getItem(excelDataArray[i].sheetName);
                        columnName = excelColumNames[excelDataArray[i].data[0].length - 1];
                        rowValue = excelDataArray[i].data.length;

                        range = sheet.getRange("A1:" + columnName + rowValue);
                        range.values = excelDataArray[i].data;
                        range.untrack();

                    }

                    return context.sync()
                        .then(function () {
                            errorLabel.innerHTML = "";
                            errorLabel.style.display = "none";
                            changElementsAvailability(actualDisableElements, false);
                            setPanelLoader("szamla-osszesito-panel-loader", "szamla-osszesito-loader", "none");
                            callback_lvl2();
                            callback_lvl1();
                        });
                });
            }
        }

        async.series(
            [
                separateWorksheets,
                clearSheets,
                addSheets,
                loadDataToSheets
            ],
            function (err) {
                console.log('all finished', err);
            }
        );
    }

    async.series(
        [
            getSzamlaOsszesito,
            workSheetHandler
        ],
        function (err) {
            console.log('allfinished', err);
        }
    )

}

function hodijAdminisztracioContainer() {
    var errorLabel = document.getElementById('hodijAdminisztracioError');
    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "block");
    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");
    //Ebben a tömbbe fognak kerülni az excelbe feltöltendő adatok
    var excelDataArray = [];

    var originalMetersResult;

    importantDisableElements = setDisableElement();
    var actualDisableElements = importantDisableElements
    changElementsAvailability(actualDisableElements, true);

    // Segédfüggvények
    var getOriginalMeters = function (callback) {

        var originalMetersCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "none");
            }
            else {
                if (result) {
                    originalMetersResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "none");
                }
            }
        }

        params = {};

        params["unit"] = "Wh";
        params["query"] = "all";
        params["type"] = "tavho";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getOriginalMeters", params, originalMetersCallback);
    }

    // Fő függvények
    var getTavhoSzerzodes = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];


        var tavhoSzerzodesCallback = function (err, tavhoSzerzodesCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "none");
            }
            else {
                if (tavhoSzerzodesCallbackResult) {
                    var requiredServerDataArray = [
                        { dataTag: "id", columnName: "A", headerText: "ID" },
                        { dataTag: "elnevezes", columnName: "B", headerText: "Elnevezés" },
                        { dataTag: "ervenyesseg_kezdete", columnName: "C", headerText: "Szerz. kezdete" },
                        { dataTag: "ervenyesseg_vege", columnName: "D", headerText: "Szerz. vége" },
                        { dataTag: "meter_identifier", columnName: "E", headerText: "Mérő azonosító" },
                        { dataTag: "alapdij", columnName: "F", headerText: "Alapdíj (Ft/hó)" },
                        { dataTag: "egysegar", columnName: "G", headerText: "Hődíj (Ft/GJ)" },

                    ];

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(tavhoSzerzodesCallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            jsonDataInnerArray.push(tavhoSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }


                    excelDataArray.push(
                        {
                            "sheetName": "Hődíj szerződések",
                            "data": jsonDataArray,
                        }
                    )

                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "none");
                }
            }
        }

        var params = {};
        params["all"] = "1";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getTavhoSzerzodes", params, tavhoSzerzodesCallback);

    }

    var getSzerzodesTovabbszamlazott = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];
        // Ebben az értékben tároljuk a mértékegységet tartalmazó oszlopok értékeit
        var unitColumnValue;

        var szerzodesTovabbszamlazottCallback = function (err, szerzodesTovabbszamlazottCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "none");
            }
            else {
                if (szerzodesTovabbszamlazottCallbackResult) {
                    var requiredServerDataArray = [
                        { dataTag: "id", columnName: "A", headerText: "ID" },
                        { dataTag: "elnevezes", columnName: "B", headerText: "Elnevezés" },
                        { dataTag: "ervenyesseg_kezdete", columnName: "C", headerText: "Érvényesség kezdete" },
                        { dataTag: "ervenyesseg_vege", columnName: "D", headerText: "Érvényesség vége" },
                        { dataTag: "tovabbado_elszamolasi_meroje_id", columnName: "E", headerText: "Továbbadó elszámolási mérője" },//originalMetersResult-ból jön
                        { dataTag: "tovabbszamlazott_mero_id", columnName: "F", headerText: "Továbbszámlázott mérő" }, // originalMetersResult-ból jön
                        { dataTag: "alapdij_felosztas", columnName: "G", headerText: "Alapdíj felosztás típusa" }, // 5 = Fajlagos, 4 = Fix költség, 3 = Almérő alapján(almérők összege), 2 = Almérő alapján (főmérőhöz), 1 = Százalékos, default = HIBA 
                        { dataTag: "alapdij_fix_koltseg_ertek", columnName: "H", headerText: "Alapdíj fix költség érték" },
                        { dataTag: "alapdij_szazalekos_ertek", columnName: "I", headerText: "Alapdíj százalékos érték" },
                        { dataTag: "almero_nelkuli", columnName: "J", headerText: "Almérő nélküli" },
                        { dataTag: "energia_felosztas", columnName: "K", headerText: "Energia felosztás típusa" }, //1= Százalékos, 2 = Almérő alapján (főmérőhöz), 3 = Almérő alapján (almérők összege), 4 = Fix költség
                        { dataTag: "energia_fix_koltseg_ertek", columnName: "L", headerText: "Energia fix költség érték" },
                        { dataTag: "energia_szazalekos_ertek", columnName: "M", headerText: "Energia százalékos érték" },
                    ];

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(szerzodesTovabbszamlazottCallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {

                            switch (requiredServerDataArray[i].dataTag) {
                                case "tovabbado_elszamolasi_meroje_id":
                                    for (var k = 0; k < originalMetersResult.length; k++) {
                                        if (originalMetersResult[k].id == szerzodesTovabbszamlazottCallbackResult.data[tmpRow].tovabbado_elszamolasi_meroje_id) {
                                            value = originalMetersResult[k].identifier;
                                            jsonDataInnerArray.push(value);
                                            break;
                                        }
                                    }
                                    break;
                                case "tovabbszamlazott_mero_id":
                                    for (var k = 0; k < originalMetersResult.length; k++) {
                                        if (originalMetersResult[k].id == szerzodesTovabbszamlazottCallbackResult.data[tmpRow].tovabbszamlazott_mero_id) {
                                            value = originalMetersResult[k].identifier;
                                            jsonDataInnerArray.push(value);
                                            break;
                                        }
                                    }
                                    break;
                                case "alapdij_felosztas":
                                    alapDijFeloszttas = szerzodesTovabbszamlazottCallbackResult.data[tmpRow].alapdij_felosztas;
                                    switch (alapDijFeloszttas) {
                                        case "1":
                                            value = "Százalékos";
                                            break;
                                        case "2":
                                            value = "Almérő alapján (főmérőhöz)";
                                            break;
                                        case "3":
                                            value = "Almérő alapján (almérők összege)";
                                            break;
                                        case "4":
                                            value = "Fix költség";
                                            break;
                                        case "5":
                                            value = "Fajlagos";
                                            break;
                                        default:
                                            value = "undefined";
                                    }
                                    jsonDataInnerArray.push(value);
                                    break;
                                case "energia_felosztas":
                                    energiaFelosztas = szerzodesTovabbszamlazottCallbackResult.data[tmpRow].energia_felosztas;
                                    switch (energiaFelosztas) {
                                        case "1":
                                            value = "Százalékos";
                                            break;
                                        case "2":
                                            value = "Almérő alapján (főmérőhöz)";
                                            break;
                                        case "3":
                                            value = "Almérő alapján (almérők összege)";
                                            break;
                                        case "4":
                                            value = "Fix költség";
                                            break;
                                        case "5":
                                            value = "Fajlagos";
                                            break;
                                        default:
                                            value = "undefined";
                                    }
                                    jsonDataInnerArray.push(value);
                                    break;


                                default:
                                    jsonDataInnerArray.push(szerzodesTovabbszamlazottCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                            }
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }

                    excelDataArray.push(
                        {
                            "sheetName": "Tovabbszámlázott szerződések",
                            "data": jsonDataArray,
                        }
                    )

                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    //Excel.run(function (context) {

                    //    var sheet = context.workbook.worksheets.getItem("Villanyszerződés Vet147");

                    //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                    //    //Excel feltöltése adatokkal
                    //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    //    range.values = jsonDataArray;
                    //    range.untrack();

                    //    // Csak a return után lesznek láthatóak az adatok az excelben
                    //    boldRange.format.font.bold = true;
                    //    return context.sync();
                    //})

                    // ---------------------EXCEL RÉSZ VÉGE --------------------
                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "none");
                }
            }
        }

        var params = {};
        params["tipus"] = "tavho";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getSzerzodesTovabbszamlazott", params, szerzodesTovabbszamlazottCallback);

    }

    //Függvény ami kezeli az excelt
    //Munkalapokat hoz létre, munkalapokat tisztít és feétölti a megfelelő munkalapokat adatokkal
    var workSheetHandler = function (callback_lvl1) {

        var clearableSheet = [];
        var addableSheet = [];

        var separateWorksheets = function (callback_lvl2) {
            // Ez a függvény a lekérdezéshez szükséges munkalapokat két külön tömbe teszi.
            // A clearableSheet tömbbe teszi a már létező munkalapok nevét
            // Az addableSheet tömbbe teszi a létrehozandó munkalapok nevét
            Excel.run(function (context) {
                var worksheets = context.workbook.worksheets;
                worksheets.load('name');
                return context.sync()
                    .then(function () {
                        var sheetFound;
                        for (var i = 0; i < excelDataArray.length; i++) {
                            sheetFound = false;
                            for (var j = 0; j < worksheets.items.length; j++) {
                                if (excelDataArray[i].sheetName == worksheets.items[j].name) {
                                    sheetFound = true;
                                    clearableSheet.push(worksheets.items[j].name);
                                    break;
                                }
                            }
                            if (sheetFound) {
                                continue;
                            }
                            else {
                                addableSheet.push(excelDataArray[i].sheetName);
                            }

                        }
                        callback_lvl2();
                    });
            })

        }

        var clearSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var addSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var loadDataToSheets = function (callback_lvl2) {
            if (excelDataArray) {
                Excel.run(function (context) {
                    var sheet;
                    var range;
                    var columnName;
                    var rowValue;

                    ////Munkalap nevének meghatározása
                    //sheet = context.workbook.worksheets.getItem("HHHCS szerződések");
                    ////Adatokkal feltöltendő tartomány meghatározésa
                    //range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    ////Adatok feltöltése
                    //range.values = jsonDataArray;
                    //range.untrack();

                    for (var i = 0; i < excelDataArray.length; i++) {
                        sheet = context.workbook.worksheets.getItem(excelDataArray[i].sheetName);
                        columnName = excelColumNames[excelDataArray[i].data[0].length - 1];
                        rowValue = excelDataArray[i].data.length;

                        range = sheet.getRange("A1:" + columnName + rowValue);
                        range.values = excelDataArray[i].data;
                        range.untrack();

                    }


                    return context.sync()
                        .then(function () {
                            errorLabel.innerHTML = "";
                            errorLabel.style.display = "none";
                            changElementsAvailability(actualDisableElements, false);
                            setPanelLoader("hodij-adminisztracio-panel-loader", "hodij-adminisztracio-loader", "none");
                            callback_lvl2();
                            callback_lvl1();
                        });
                });
            }
        }

        async.series(
            [
                separateWorksheets,
                clearSheets,
                addSheets,
                loadDataToSheets
            ],
            function (err) {
                console.log('all finished', err);
            }
        );
    }

    async.series(
        [
            getOriginalMeters,
            getTavhoSzerzodes,
            getSzerzodesTovabbszamlazott,
            workSheetHandler
        ],
        function (err) {
            console.log('allfinished', err);
        }
    )


}

function szakreferensiJelentesContainer() {
    var errorLabel = document.getElementById('szakreferensiJelentesError');
    var maxRequestbeginDate = new Date();
    var currentYear = maxRequestbeginDate.getFullYear();

    var inputFullContent = document.getElementById('szakreferensiJelentesYearFilter').value;
    var inputYear = document.getElementById('szakreferensiJelentesYearFilter').value.substring(0, 4);

    // Input megfelelőségi és Regex teszetek
    if (isNaN(inputYear) == true) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A megadott év nem megfelelő."
        return;
    }

    if (inputYear > currentYear) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A megadott év a jövőben van."
        return;
    }

    if (/^([12]\d{3}-(0[1-9]|1[0-2]))$/
        .test(inputFullContent) == false) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A tárgy időszak dátum nem megfelelő formátumú. Helyes formátum (YYYY-MM)"
        return;
    }

    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "block");

    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");
    //Ebben a tömbbe fognak kerülni az excelbe feltöltendő adatok
    var excelDataArray = []

    //var threadLimit;

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    importantDisableElements = setDisableElement();
    var newDisableElements = ["szakreferensiJelentesYearFilter", "szakreferensi_jelentes_meter_groups", "szakreferensi_jelentes_mentett_bealitasok"];
    var actualDisableElements = newDisableElements.concat(importantDisableElements);
    changElementsAvailability(actualDisableElements, true);

    // Szükséges változók
    var savedOptionsList = document.getElementById('szakreferensi_jelentes_mentett_bealitasok');
    var savedOptionsListSelectedText;
    try {
        savedOptionsListSelectedText = savedOptionsList.options[savedOptionsList.selectedIndex].text;
    } catch (e) {
        savedOptionsListSelectedText = ""
    }

    var meterGroupArrayResult;
    var savedOptionsArray;
    var meterTreeArray;

    // Segédfüggvények

    var meterGroup = function (callback) {

        var getMeterGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
            }
            else {
                if (result) {
                    meterGroupArrayResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
                }
            }
        }

        params = {};

        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, getMeterGroupCallback);

    }

    var getSavedGraphs = function (callback) {
        var savedGraphsCallBack = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
            }
            else {
                if (result) {
                    savedOptionsArray = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
                }
            }
        }

        var params = {};
        params["is_public"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/mdgraph/draw/getSavedGraphs", params, savedGraphsCallBack);

    }

    var getMeterTree = function (callback) {

        var metreGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
            }
            else {
                if (result) {
                    meterTreeArray = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
                }
            }
        }

        var params = {};

        params["node"] = "";
        params["page"] = "";

        postAsyncGetData(host + "/mdgraph/draw/getMeterTree", params, metreGroupCallback);

    }

    // Fő függvények
    var getFeldolgozottMeresek = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var feldolgozottMeresekDateArray = [];

        var feldolgozottMeresekGetDate = parseInt(document.getElementById('szakreferensiJelentesYearFilter').value.substring(0, 4));
        for (var i = 0; i <= 3; i++) {
            feldolgozottMeresekDateArray.push((feldolgozottMeresekGetDate-i));
        }

        var meterGroupList = document.getElementById('szakreferensi_jelentes_meter_groups');
        var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;
        // A szerverlekérdezéshez szükséges 'meter_gruop' paraméter értékét meghatározó változó
        var meterGroupValue;

        for (var i = 0; i < meterGroupArrayResult.length; i++) {
            if (meterGroupArrayResult[i].nev == meterGroupListSelectedText) {
                meterGroupValue = meterGroupArrayResult[i].id;
                break;
            }
        }

        var params = {};
        params["meter_group"] = meterGroupValue;
        params["napok_mutatasa"] = "false";
        params["calculated_natural_gas"] = "0";
        params["tankolas_is"] = "1";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";



        var feldolgozottMeresekValues = function (item, innerCallback) {

            var feldolgozottMeresekValuesCallback = function (err, feldolgozottMeresekValuesCallbackResult) {

                var requiredServerDataArray = [
                    { dataTag: "identifier", columnName: "A", headerText: "Mérő azonosító" },
                    { dataTag: "name", columnName: "B", headerText: "Megnevezés" },
                    { dataTag: "ho1", columnName: "C", headerText: item + "-01" },
                    { dataTag: "ho2", columnName: "D", headerText: item + "-02" },
                    { dataTag: "ho3", columnName: "E", headerText: item + "-03" },
                    { dataTag: "ho4", columnName: "F", headerText: item + "-04" },
                    { dataTag: "ho5", columnName: "G", headerText: item + "-05" },
                    { dataTag: "ho6", columnName: "H", headerText: item + "-06" },
                    { dataTag: "ho7", columnName: "I", headerText: item + "-07" },
                    { dataTag: "ho8", columnName: "J", headerText: item + "-08" },
                    { dataTag: "ho9", columnName: "K", headerText: item + "-09" },
                    { dataTag: "ho10", columnName: "L", headerText: item + "-10" },
                    { dataTag: "ho11", columnName: "M", headerText: item + "-11" },
                    { dataTag: "ho12", columnName: "N", headerText: item + "-12" },
                ];

                //Fejlécek betöltése a jsonDataArray-ba
                jsonDataInnerArray = [];
                jsonDataArray = [];
                requiredServerDataArray.forEach(function (element) {
                    jsonDataInnerArray.push(element.headerText);
                });
                jsonDataArray.push(jsonDataInnerArray);
                jsonDataInnerArray = [];

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = feldolgozottMeresekValuesCallbackResult.data.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {

                                jsonDataInnerArray.push(feldolgozottMeresekValuesCallbackResult.data[i][requiredServerDataArray[j].dataTag]);

                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }

                    var workSheetPrefixNumber = (feldolgozottMeresekGetDate - item);
                    excelDataArray.push(
                        {
                            "sheetName": "IN_É" + workSheetPrefixNumber,
                            "data": jsonDataArray,
                        }
                    )

                    requestCounter++;
                    if (requestCounter == feldolgozottMeresekDateArray.length) {
                        //if (item == katPenzeszkozokArray[katPenzeszkozokArray.length - 1]) {
                        dataLength = jsonDataArray.length - 1;

                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["datum_meres_kezdete"] = item + "-01";
            postAsyncGetData(host + "/ebill/summary/getFeldolgozottMeresek", params, feldolgozottMeresekValuesCallback);
        };

        async.eachLimit(
            feldolgozottMeresekDateArray,
            threadLimit,
            feldolgozottMeresekValues,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getMentettBeallitasokGrafikonAdatok = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var mentettBeallitasokGrafikonAdatokCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
            }
            else {
                if (result) {
                    var extraInfoObj = result.extraInfo;
                    var extraInfoKeysArray = [];
                    for (var k in extraInfoObj) extraInfoKeysArray.push(k.replace("value", ""));

                    dataLength = Object.keys(meterTreeArray.data).length;
                    var headerArray = [];

                    extraInfoKeysArray.forEach(function (element) {

                        for (var i = 0; i < dataLength; i++) {
                            var dataSecondLevelLength = Object.keys(meterTreeArray.data[i].data).length
                            for (var j = 0; j < dataSecondLevelLength; j++) {
                                if (meterTreeArray.data[i].data[j].meter_id == element) {

                                    var elementHeaderCompatibleString = "value" + element;
                                    headerArray.push({ extraInfoKey: elementHeaderCompatibleString, extraInfoText: meterTreeArray.data[i].data[j].identifier });
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

                    jsonDataArray = [];
                    jsonDataInnerArray = [];

                    //Fejlécek betöltése a jsonDataArray-ba
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(result.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    var correctDateWithFormat;
                    var isHour = false;
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {
                            if (requiredServerDataArray[i].dataTag == "tstamp") {
                                var d = new Date(result.data[tmpRow][requiredServerDataArray[i].dataTag]);

                                correctDateWithFormat = d.getFullYear().toString() + "-" + ((d.getMonth() + 1).toString().length == 2 ? (d.getMonth() + 1).toString() : "0" + (d.getMonth() + 1).toString()) + "-" + (d.getDate().toString().length == 2 ? d.getDate().toString() : "0" + d.getDate().toString()) + " " + (d.getHours().toString().length == 2 ? d.getHours().toString() : " " + d.getHours().toString()) + ":" + ((parseInt(d.getMinutes() / 5) * 5).toString().length == 2 ? (parseInt(d.getMinutes() / 5) * 5).toString() : "0" + (parseInt(d.getMinutes() / 5) * 5).toString()) + ":00";

                                goddamn = (correctDateWithFormat.substring(correctDateWithFormat.length - 5, correctDateWithFormat.length - 3));
                                if ((correctDateWithFormat.substring(correctDateWithFormat.length - 5, correctDateWithFormat.length - 3)) != "00") {
                                    isHour = false;
                                    break;
                                    
                                }
                                else {
                                    isHour = true;
                                }
                                jsonDataInnerArray.push(correctDateWithFormat);

                            }
                            else {
                                jsonDataInnerArray.push(result.data[tmpRow][requiredServerDataArray[i].dataTag]);
                            }
                        }

                        if (isHour) {
                            jsonDataArray.push(jsonDataInnerArray);
                            isHour = false;
                        } else {
                            continue;
                        }
                        
                    }

                    excelDataArray.push(
                        {
                            "sheetName": "IN_D0",
                            "data": jsonDataArray,
                        }
                    )
                    //Caolan async miatt
                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás";
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("heti-jelentes-panel-loader", "heti-jelentes-loader", "none");
                }
            }
        }

        var datetime_from;
        var datetime_to;
        if (document.getElementById("szakreferensi_jelentes_eves_lekerdezes_checkbox").checked == false) {
            //A datetime_from változó lesz a getGraphSeries lekérdezés datetime_from paramétere
            datetime_from = inputFullContent + "-01;00:00";

            //A datetime_to változó lesz a getGraphSeries lekérdezés datetime_to paramétere
            let monthCheck = parseInt(inputFullContent.substring(5, 7));
            if (monthCheck == 12) {
                datetime_to = (parseInt(inputFullContent.substring(0, 4)) + 1) + "-01-01;06:00"
            }
            else {
                if (monthCheck < 9) {
                    datetime_to = inputFullContent.substring(0, 4) + "-0" + (monthCheck + 1) + "-01;06:00"
                } else {
                    datetime_to = inputFullContent.substring(0, 4) + "-" + (monthCheck + 1) + "-01;06:00"
                }
            }
        }
        else {
            datetime_from = inputYear + "-01-01;00:00";
            datetime_to = inputYear + "-12-01;06:00";
        }




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

        dataLength = Object.keys(meterTreeArray.data).length

        var type_list_string = "";

        savedOptionsMetersArray.forEach(function (element) {

            for (var i = 0; i < dataLength; i++) {
                var dataSecondLevelLength = Object.keys(meterTreeArray.data[i].data).length
                for (var j = 0; j < dataSecondLevelLength; j++) {
                    if (meterTreeArray.data[i].data[j].meter_id == element) {

                        type_list_string = type_list_string.concat(meterTreeArray.data[i].data[j].data_type_id, ",");
                        i = dataLength;
                        break;
                    }
                }
            }
        });
        //A type_list_string változó lesz getGraphSeries lekérdezés type_list paramétere
        type_list_string = type_list_string.slice(0, -1);

        var params = {};

        params["datetime_from"] = datetime_from;
        params["datetime_to"] = datetime_to;
        params["meter_list"] = savedOptionsMeters;
        params["baseline_list"] = "";
        params["type_list"] = type_list_string;
        params["serie_type"] = "11";
        params["resolution"] = savedOptionsResolution;
        params["type"] = savedOptionsType;
        params["sendTo"] = "";
        params["checker"] = "0";
        params["extraInfo"] = "1";
        params["fake"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "9999999";

        postAsyncGetData(host + "/mdgraph/draw/getGraphSeries", params, mentettBeallitasokGrafikonAdatokCallback);
    }

    var getFogyasztasOsszesito = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var threadLimit = 2;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var fogyasztasOsszesitoDateArray = [];

        if (document.getElementById("szakreferensi_jelentes_eves_lekerdezes_checkbox").checked == false) {
            fogyasztasOsszesitoDateArray.push(inputFullContent);
        }
        else {
            for (var i = 1; i <= 12; i++) {
                if (i.toString().length == 1) {
                    monthValue = "0" + i;
                } else {
                    monthValue = i;
                }
                fogyasztasOsszesitoDateArray.push(inputYear + "-" + monthValue);
            }
        }

        var meterGroupList = document.getElementById('szakreferensi_jelentes_meter_groups');
        var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;
        // A szerverlekérdezéshez szükséges 'meter_gruop' paraméter értékét meghatározó változó
        var meterGroupValue;

        for (var i = 0; i < meterGroupArrayResult.length; i++) {
            if (meterGroupArrayResult[i].nev == meterGroupListSelectedText) {
                meterGroupValue = meterGroupArrayResult[i].id;
                break;
            }
        }

        var params = {};
        params["meter_group"] = meterGroupValue;
        params["sendTo"] = "screen";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";



        var fogyasztasOsszesitoValues = function (item, innerCallback) {
            var workSheetPrefixNumber = parseInt(item.substring(item.length - 2, item.length));

            var fogyasztasOsszesitoValuesCallback = function (err, fogyasztasOsszesitoValuesCallbackResult) {

                var requiredServerDataArray = [
                    { dataTag: "meter_name", columnName: "A", headerText: "Mérés neve" },
                    { dataTag: "meter_identifier", columnName: "B", headerText: "Mérő azonosító" },
                    { dataTag: "pod_azonosito", columnName: "C", headerText: "POD" },
                    { dataTag: "idoszak_kezdete", columnName: "D", headerText: "Időszak kezdete" },
                    { dataTag: "idoszak_vege", columnName: "E", headerText: "Időszak vége" },
                    { dataTag: "tarifa_hosszu_nev", columnName: "F", headerText: "Tarifa" },
                    { dataTag: "lekotott_teljesitmeny", columnName: "G", headerText: "Lekötött telj." },
                    { dataTag: "lekotott_teljesitmeny_mertekegyseg", columnName: "H", headerText: "[]" },
                    { dataTag: "operativ_teljesitmeny", columnName: "I", headerText: "Operatív teljesítmény" },
                    { dataTag: "operativ_teljesitmeny_mertekegyseg", columnName: "J", headerText: "[]" },
                    { dataTag: "max_teljesitmeny", columnName: "K", headerText: "Max. telj." },
                    { dataTag: "max_teljesitmeny_mertekegyseg", columnName: "L", headerText: "[]" },
                    { dataTag: "fogyasztas", columnName: "M", headerText: "Fogyasztás" },
                    { dataTag: "fogyasztas_mertekegyseg", columnName: "N", headerText: "[]" },
                    { dataTag: "fogyasztas_elozo_ev", columnName: "O", headerText: "Előző évi fogyasztás" },
                    { dataTag: "fogyasztas_elozo_ev_mertekegyseg", columnName: "P", headerText: "[]" },
                    { dataTag: "induktiv_tulfogyasztas", columnName: "Q", headerText: "Induktív túl fogy." },
                    { dataTag: "induktiv_tulfogyasztas_mertekegyseg", columnName: "R", headerText: "[]" },
                    { dataTag: "kapacitiv_fogyasztas", columnName: "S", headerText: "Kapacitív fogy." },
                    { dataTag: "kapacitiv_fogyasztas_mertekegyseg", columnName: "T", headerText: "[]" },
                ];

                //Fejlécek betöltése a jsonDataArray-ba
                jsonDataInnerArray = [];
                jsonDataArray = [];
                requiredServerDataArray.forEach(function (element) {
                    jsonDataInnerArray.push(element.headerText);
                });
                jsonDataArray.push(jsonDataInnerArray);
                jsonDataInnerArray = [];

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = fogyasztasOsszesitoValuesCallbackResult.data.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                if (requiredServerDataArray[j].dataTag == "idoszak_kezdete" || requiredServerDataArray[j].dataTag == "idoszak_vege") {
                                    jsonDataInnerArray.push("'" + ((fogyasztasOsszesitoValuesCallbackResult.data[i][requiredServerDataArray[j].dataTag]).replace(/\./g, "-")));
                                } else {
                                    jsonDataInnerArray.push(fogyasztasOsszesitoValuesCallbackResult.data[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }

                    //var workSheetPrefixNumber = parseInt(item.substring(item.length - 2, item.length));



                    excelDataArray.push(
                        {
                            "sheetName": "IN_F" + workSheetPrefixNumber,
                            "data": jsonDataArray,
                        }
                    )

                    requestCounter++;
                    if (requestCounter == fogyasztasOsszesitoDateArray.length) {
                        //if (item == katPenzeszkozokArray[katPenzeszkozokArray.length - 1]) {
                        dataLength = jsonDataArray.length - 1;

                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            if (workSheetPrefixNumber.toString().length == 1 && workSheetPrefixNumber != 9) {
                params["date_to"] = item.substring(0, 4) + "0" + (workSheetPrefixNumber + 1) + "-01"; //T00:00:00
            } else {
                params["date_to"] = item.substring(0, 4) + (workSheetPrefixNumber + 1) + "-01"; //T00:00:00
            }

            switch (true) {
                case (workSheetPrefixNumber == 12):
                    params["date_to"] = "" + (parseInt(item.substring(0, 4)) + 1) + "-01-01";//T00:00:00
                    break;
                case (workSheetPrefixNumber >= 9):
                    params["date_to"] = "" + item.substring(0, 4) + "-" + (workSheetPrefixNumber + 1) + "-01"//T00:00:00
                    break;
                case (workSheetPrefixNumber > 0 || workSheetPrefixNumber < 9):
                    params["date_to"] = "" + item.substring(0, 4) + "-0" + (workSheetPrefixNumber + 1) + "-01"; //T00:00:00
                    break;
                default: console.log("There are problems with workSheetPrefixNumber in fogyasztasOsszesitoValues function");
            }

            params["date_from"] = item + "-01";//T00:00:00
            postAsyncGetData(host + "/ebill/billing/getFogyasztasOsszesito2", params, fogyasztasOsszesitoValuesCallback);
        };

        async.eachLimit(
            fogyasztasOsszesitoDateArray,
            threadLimit,
            fogyasztasOsszesitoValues,
            function (err) {
                console.log('all finished', err);
            }
        );

    }
    
    //Függvény ami kezeli az excelt
    //Munkalapokat hoz létre, munkalapokat tisztít és feétölti a megfelelő munkalapokat adatokkal
    var workSheetHandler = function (callback_lvl1) {

        var clearableSheet = [];
        var addableSheet = [];

        var separateWorksheets = function (callback_lvl2) {
            // Ez a függvény a lekérdezéshez szükséges munkalapokat két külön tömbe teszi.
            // A clearableSheet tömbbe teszi a már létező munkalapok nevét
            // Az addableSheet tömbbe teszi a létrehozandó munkalapok nevét
            Excel.run(function (context) {
                var worksheets = context.workbook.worksheets;
                worksheets.load('name');
                return context.sync()
                    .then(function () {
                        var sheetFound;
                        for (var i = 0; i < excelDataArray.length; i++) {
                            sheetFound = false;
                            for (var j = 0; j < worksheets.items.length; j++) {
                                if (excelDataArray[i].sheetName == worksheets.items[j].name) {
                                    sheetFound = true;
                                    clearableSheet.push(worksheets.items[j].name);
                                    break;
                                }
                            }
                            if (sheetFound) {
                                continue;
                            }
                            else {
                                addableSheet.push(excelDataArray[i].sheetName);
                            }

                        }
                        callback_lvl2();
                    });
            })

        }

        var clearSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var addSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var loadDataToSheets = function (callback_lvl2) {
            if (excelDataArray) {
                Excel.run(function (context) {
                    var sheet;
                    var range;
                    var columnName;
                    var rowValue;

                    ////Munkalap nevének meghatározása
                    //sheet = context.workbook.worksheets.getItem("HHHCS szerződések");
                    ////Adatokkal feltöltendő tartomány meghatározésa
                    //range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    ////Adatok feltöltése
                    //range.values = jsonDataArray;
                    //range.untrack();

                    for (var i = 0; i < excelDataArray.length; i++) {
                        sheet = context.workbook.worksheets.getItem(excelDataArray[i].sheetName);
                        columnName = excelColumNames[excelDataArray[i].data[0].length - 1];
                        rowValue = excelDataArray[i].data.length;

                        range = sheet.getRange("A1:" + columnName + rowValue);
                        range.values = excelDataArray[i].data;
                        range.untrack();

                    }


                    return context.sync()
                        .then(function () {
                            errorLabel.innerHTML = "";
                            errorLabel.style.display = "none";
                            changElementsAvailability(actualDisableElements, false);
                            setPanelLoader("szakreferensi-jelentes-panel-loader", "szakreferensi-jelentes-loader", "none");
                            callback_lvl2();
                            callback_lvl1();
                        });
                });
            }
        }

        async.series(
            [
                separateWorksheets,
                clearSheets,
                addSheets,
                loadDataToSheets
            ],
            function (err) {
                console.log('all finished', err);
            }
        );
    }

    var asyncSeriesFunctionsArray = [
        meterGroup,
        getSavedGraphs,
        getMeterTree,
        getFeldolgozottMeresek,
        getMentettBeallitasokGrafikonAdatok,
        getFogyasztasOsszesito,
        workSheetHandler
    ];

    if (document.getElementById('szakreferensi_jelentes_meter_groups').options.length == 0) {
        for (var i = 0; i < asyncSeriesFunctionsArray.length; i++) {
            if (asyncSeriesFunctionsArray[i] == getFeldolgozottMeresek) {
                asyncSeriesFunctionsArray.splice(i, 1);
                break;
            }
        }
    }

    if (document.getElementById('szakreferensi_jelentes_mentett_bealitasok').options.length == 0) {
        for (var i = 0; i < asyncSeriesFunctionsArray.length; i++) {
            if (asyncSeriesFunctionsArray[i] == getMentettBeallitasokGrafikonAdatok) {
                asyncSeriesFunctionsArray.splice(i, 1);
                break;
            }
        }
    }

    async.series(
        asyncSeriesFunctionsArray,
        function (err) {
            console.log('allfinished', err);
        }
    )
}

function rezsiCsokkentesContainer() {
    var errorLabel = document.getElementById('rezsicsokkentesError');
    var maxRequestbeginDate = new Date();
    var currentYear = maxRequestbeginDate.getFullYear();

    var inputFullContent = document.getElementById('rezsiCsokkentesYearFilter').value;
    var inputYear = document.getElementById('rezsiCsokkentesYearFilter').value.substring(0, 4);
    var inputMonth = parseInt(inputFullContent.substring(inputFullContent.length - 2, inputFullContent.length));

    // Input megfelelőségi és Regex teszetek
    if (isNaN(inputYear) == true) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A megadott év nem megfelelő."
        return;
    }

    if (inputYear > currentYear) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A megadott év a jövőben van."
        return;
    }

    if (/^([12]\d{3}-(0[1-9]|1[0-2]))$/
        .test(inputFullContent) == false) {
        errorLabel.style.display = 'block';
        errorLabel.innerHTML = "A tárgy időszak dátum nem megfelelő formátumú. Helyes formátum (YYYY-MM)"
        return;
    }

    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';
    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "block");

    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");
    //Ebben a tömbbe fognak kerülni az excelbe feltöltendő adatok
    var excelDataArray = []

    //var threadLimit;

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    importantDisableElements = setDisableElement();
    var newDisableElements = ["rezsiCsokkentesYearFilter", "rezsi_csokkentes_meter_groups"];
    var actualDisableElements = newDisableElements.concat(importantDisableElements);
    changElementsAvailability(actualDisableElements, true);

    //Segédváltozók
    var meterGroupArrayResult;
    var csatlakozasiPontResult;
    var rhdDatumok;
    var rhdValues = [];
    var meddoWattosDatumok;
    var HHHCSIds = [];
    var tmpCounter = 2;



    // Segédfüggvények

    var meterGroup = function (callback) {

        var getMeterGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
            }
            else {
                if (result) {
                    meterGroupArrayResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
            }
        }

        params = {};

        params["query"] = "all";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, getMeterGroupCallback);

    }

    var getCsatlakozasiPont = function (callback) {

        var getCsatlakozasiPontCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
            }
            else {
                if (result) {
                    csatlakozasiPontResult = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
            }
        }

        params = {};

        params["all"] = "1";
        params["isMasodlagos"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getHCSSzerzodes", params, getCsatlakozasiPontCallback);

    }

    var getMeddoWattosAranyValidFrom = function (callback) {

        var getMeddoWattosAranyValidFromCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
            }
            else {
                if (result) {
                    meddoWattosDatumok = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
            }
        }

        params = {};

        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getMeddoWattosAranyValidFrom", params, getMeddoWattosAranyValidFromCallback);
    }

    var getRHDValidFrom = function (callback) {

        var getRHDValidFromCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
            }
            else {
                if (result) {
                    rhdDatumok = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
            }
        }

        params = {};

        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/admin/getRHDValidFrom", params, getRHDValidFromCallback);
    }

    var getAllandoRendszerhasznalatiDijak = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var rhdDatumokArray = [];
        for (var i = 0; i < rhdDatumok.length; i++) {
            rhdDatumokArray.push(rhdDatumok[i].ervenyesseg_kezdete);
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "ervenyesseg_kezdete", columnName: "A", headerText: "Érvényesség kezdete" },
            { dataTag: "ervenyesseg_vege", columnName: "B", headerText: "Érvényesség vége" },
            { dataTag: "atviteli_rendszeriranyitasi_dij", columnName: "C", headerText: "Átviteli rendszerirányítási díj (Ft/kWh)" },
            { dataTag: "rendszerszintu_szolgaltatasi_dij", columnName: "D", headerText: "Rendszerszintű szolgáltatási díj (Ft/kWh)" },
            { dataTag: "kozvilagitasi_elosztasi_dij", columnName: "E", headerText: "Közvilágítási elosztási díj (Ft/kWh)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var AllandoRendszerhasznalatiDijak = function (item, innerCallback) {

            var AllandoRendszerhasznalatiDijakCallback = function (err, AllandoRendszerhasznalatiDijakCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
                else {
                    //RHD adatok kimetése, hogy a következő fő függvény is tudja használni
                    // EZ IS FUNKCIONÁL MINT SEGÉDFÜGGVÉNY!!!!
                    rhdValues.push(AllandoRendszerhasznalatiDijakCallbackResult);

                    dataLength = AllandoRendszerhasznalatiDijakCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {

                                switch (requiredServerDataArray[j].dataTag) {
                                    case "atviteli_rendszeriranyitasi_dij":
                                        value = AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "kozvilagitasi_elosztasi_dij":
                                        value = AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "rendszerszintu_szolgaltatasi_dij":
                                        value = AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    default:
                                        jsonDataInnerArray.push(AllandoRendszerhasznalatiDijakCallbackResult[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }
                    requestCounter++;

                    //if (item == rhdDatumok[rhdDatumok.length - 1].ervenyesseg_kezdete) {
                    if (requestCounter == rhdDatumok.length) {
                        dataLength = jsonDataArray.length - 1;

                        //excelDataArray.push(
                        //    {
                        //        "sheetName": "RHD azonos",
                        //        "data": jsonDataArray,
                        //    }
                        //)


                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("RHD azonos");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item;
            postAsyncGetData(host + "/ebill/admin/getRHDAllandoValues_mertekegyseggel", params, AllandoRendszerhasznalatiDijakCallback);
        };

        async.eachLimit(
            rhdDatumokArray,
            threadLimit,
            AllandoRendszerhasznalatiDijak,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    //Főfüggvények

    var getFogyasztasOsszesito = function (callback) {
        //----------------------------------------------------------
        //MEGJEGYZÉS: Ez tákolt függvény ne használd másolásra!!!
        //----------------------------------------------------------

        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        var threadLimit = 4;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var fogyasztasOsszesitoDateArray = [];

            for (var i = (inputMonth - 2); i <= inputMonth; i++) {
                switch (true) {
                    case (i < 1):
                        fogyasztasOsszesitoDateArray.push({ "date": (parseInt(inputYear) - 1) + "-" + (12 + i), "counter": tmpCounter});
                        break;
                    default:
                        if (i.toString().length == 1) {
                            fogyasztasOsszesitoDateArray.push({ "date": inputYear + "-0" + i, "counter": tmpCounter})
                        } else {
                            fogyasztasOsszesitoDateArray.push({ "date": inputYear + "-" + i, "counter": tmpCounter })
                        }
                        
                }
                tmpCounter--;
        }

        var meterGroupList = document.getElementById('rezsi_csokkentes_meter_groups');
        var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;
        // A szerverlekérdezéshez szükséges 'meter_gruop' paraméter értékét meghatározó változó
        var meterGroupValue;

        for (var i = 0; i < meterGroupArrayResult.length; i++) {
            if (meterGroupArrayResult[i].nev == meterGroupListSelectedText) {
                meterGroupValue = meterGroupArrayResult[i].id;
                break;
            }
        }

        var params = {};
        params["meter_group"] = meterGroupValue;
        params["sendTo"] = "screen";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";



        var fogyasztasOsszesitoValues = function (item, innerCallback) {
            var workSheetPrefixNumber = parseInt(item.date.substring(item.date.length - 2, item.date.length));

            var fogyasztasOsszesitoValuesCallback = function (err, fogyasztasOsszesitoValuesCallbackResult) {

                var requiredServerDataArray = [
                    { dataTag: "meter_name", columnName: "A", headerText: "Mérés neve" },
                    { dataTag: "meter_identifier", columnName: "B", headerText: "Mérő azonosító" },
                    { dataTag: "pod_azonosito", columnName: "C", headerText: "POD" },
                    { dataTag: "idoszak_kezdete", columnName: "D", headerText: "Időszak kezdete" },
                    { dataTag: "idoszak_vege", columnName: "E", headerText: "Időszak vége" },
                    { dataTag: "tarifa_hosszu_nev", columnName: "F", headerText: "Tarifa" },
                    { dataTag: "lekotott_teljesitmeny", columnName: "G", headerText: "Lekötött telj." },
                    { dataTag: "lekotott_teljesitmeny_mertekegyseg", columnName: "H", headerText: "[]" },
                    { dataTag: "operativ_teljesitmeny", columnName: "I", headerText: "Operatív teljesítmény" },
                    { dataTag: "operativ_teljesitmeny_mertekegyseg", columnName: "J", headerText: "[]" },
                    { dataTag: "max_teljesitmeny", columnName: "K", headerText: "Max. telj." },
                    { dataTag: "max_teljesitmeny_mertekegyseg", columnName: "L", headerText: "[]" },
                    { dataTag: "fogyasztas", columnName: "M", headerText: "Fogyasztás" },
                    { dataTag: "fogyasztas_mertekegyseg", columnName: "N", headerText: "[]" },
                    { dataTag: "fogyasztas_elozo_ev", columnName: "O", headerText: "Előző évi fogyasztás" },
                    { dataTag: "fogyasztas_elozo_ev_mertekegyseg", columnName: "P", headerText: "[]" },
                    { dataTag: "induktiv_tulfogyasztas", columnName: "Q", headerText: "Induktív túl fogy." },
                    { dataTag: "induktiv_tulfogyasztas_mertekegyseg", columnName: "R", headerText: "[]" },
                    { dataTag: "kapacitiv_fogyasztas", columnName: "S", headerText: "Kapacitív fogy." },
                    { dataTag: "kapacitiv_fogyasztas_mertekegyseg", columnName: "T", headerText: "[]" },
                ];

                //Fejlécek betöltése a jsonDataArray-ba
                jsonDataInnerArray = [];
                jsonDataArray = [];
                requiredServerDataArray.forEach(function (element) {
                    jsonDataInnerArray.push(element.headerText);
                });
                jsonDataArray.push(jsonDataInnerArray);
                jsonDataInnerArray = [];

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = fogyasztasOsszesitoValuesCallbackResult.data.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                if (requiredServerDataArray[j].dataTag == "idoszak_kezdete" || requiredServerDataArray[j].dataTag == "idoszak_vege") {
                                    jsonDataInnerArray.push("'" + ((fogyasztasOsszesitoValuesCallbackResult.data[i][requiredServerDataArray[j].dataTag]).replace(/\./g, "-")));
                                } else {
                                    jsonDataInnerArray.push(fogyasztasOsszesitoValuesCallbackResult.data[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }

                    //var workSheetPrefixNumber = parseInt(item.substring(item.length - 2, item.length));



                    excelDataArray.push(
                        {
                            "sheetName": "IN_F" + item.counter,
                            "data": jsonDataArray,
                        }
                    )

                    requestCounter++;
                    if (requestCounter == fogyasztasOsszesitoDateArray.length) {
                        //if (item == katPenzeszkozokArray[katPenzeszkozokArray.length - 1]) {
                        dataLength = jsonDataArray.length - 1;

                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            if (workSheetPrefixNumber.toString().length == 1 && workSheetPrefixNumber != 9) {
                params["date_to"] = item.date.substring(0, 4) + "0" + (workSheetPrefixNumber + 1) + "-01"; //T00:00:00
            } else {
                params["date_to"] = item.date.substring(0, 4) + (workSheetPrefixNumber + 1) + "-01"; //T00:00:00
            }

            switch (true) {
                case (workSheetPrefixNumber == 12):
                    params["date_to"] = "" + (parseInt(item.date.substring(0, 4)) + 1) + "-01-01";//T00:00:00
                    break;
                case (workSheetPrefixNumber >= 9):
                    params["date_to"] = "" + item.date.substring(0, 4) + "-" + (workSheetPrefixNumber + 1) + "-01"//T00:00:00
                    break;
                case (workSheetPrefixNumber > 0 || workSheetPrefixNumber < 9):
                    params["date_to"] = "" + item.date.substring(0, 4) + "-0" + (workSheetPrefixNumber + 1) + "-01"; //T00:00:00
                    break;
                default: console.log("There are problems with workSheetPrefixNumber in fogyasztasOsszesitoValues function");
            }

            params["date_from"] = item.date + "-01";//T00:00:00
            postAsyncGetData(host + "/ebill/billing/getFogyasztasOsszesito2", params, fogyasztasOsszesitoValuesCallback);
        };

        async.eachLimit(
            fogyasztasOsszesitoDateArray,
            threadLimit,
            fogyasztasOsszesitoValues,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getFeldolgozottMeresek = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var feldolgozottMeresekDateArray = [];

        var feldolgozottMeresekGetDate = parseInt(document.getElementById('rezsiCsokkentesYearFilter').value.substring(0, 4));
        //for (var i = 0; i <= 3; i++) {
        //    feldolgozottMeresekDateArray.push((feldolgozottMeresekGetDate - i));
        //}

        feldolgozottMeresekDateArray.push(feldolgozottMeresekGetDate);

        var meterGroupList = document.getElementById('rezsi_csokkentes_meter_groups');
        var meterGroupListSelectedText = meterGroupList.options[meterGroupList.selectedIndex].text;
        // A szerverlekérdezéshez szükséges 'meter_gruop' paraméter értékét meghatározó változó
        var meterGroupValue;

        for (var i = 0; i < meterGroupArrayResult.length; i++) {
            if (meterGroupArrayResult[i].nev == meterGroupListSelectedText) {
                meterGroupValue = meterGroupArrayResult[i].id;
                break;
            }
        }

        var params = {};
        params["meter_group"] = meterGroupValue;
        params["napok_mutatasa"] = "true";
        params["calculated_natural_gas"] = "0";
        params["tankolas_is"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";



        var feldolgozottMeresekValues = function (item, innerCallback) {

            var feldolgozottMeresekValuesCallback = function (err, feldolgozottMeresekValuesCallbackResult) {

                var requiredServerDataArray = [
                    { dataTag: "identifier", columnName: "A", headerText: "Mérő azonosító" },
                    { dataTag: "name", columnName: "B", headerText: "Megnevezése" },
                    { dataTag: "type_name", columnName: "C", headerText: "Típus megnevezése" },
                ];

                let actMonthText;
                let tmpExcelColumn = 3
                for (var i = 1; i <= 12; i++) {
                    if (i.toString().length == 1) {
                        actMonthText = '0' + i;
                    } else {
                        actMonthText = i;
                    }
                    requiredServerDataArray.push({ dataTag: "ho" + i, columnName: excelColumNames[tmpExcelColumn], headerText: item + "-" + actMonthText })
                    tmpExcelColumn++;
                    requiredServerDataArray.push({ dataTag: "ho" + i + "_day_count", columnName: excelColumNames[tmpExcelColumn], headerText: i + " Hónap : adatok" })
                    tmpExcelColumn++;
                }

                //Fejlécek betöltése a jsonDataArray-ba
                jsonDataInnerArray = [];
                jsonDataArray = [];
                requiredServerDataArray.forEach(function (element) {
                    jsonDataInnerArray.push(element.headerText);
                });
                jsonDataArray.push(jsonDataInnerArray);
                jsonDataInnerArray = [];

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = feldolgozottMeresekValuesCallbackResult.data.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                if (requiredServerDataArray[j].dataTag.includes("_day_count")) {
                                    repairString = feldolgozottMeresekValuesCallbackResult.data[i][requiredServerDataArray[j].dataTag]
                                    //ITT KELL KIJAVITANI A REPAIRSTRINGET
                                    pos1 = repairString.indexOf(">") + 1;
                                    pos2 = repairString.indexOf("<", 2);
                                    repairString = repairString.substring(pos1, pos2);
                                    jsonDataInnerArray.push(repairString);
                                } else {
                                    jsonDataInnerArray.push(feldolgozottMeresekValuesCallbackResult.data[i][requiredServerDataArray[j].dataTag]);
                                }
                                

                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }

                    var workSheetPrefixNumber = (feldolgozottMeresekGetDate - item);
                    excelDataArray.push(
                        {
                            "sheetName": "IN_É" + workSheetPrefixNumber,
                            "data": jsonDataArray,
                        }
                    )

                    requestCounter++;
                    if (requestCounter == feldolgozottMeresekDateArray.length) {
                        //if (item == katPenzeszkozokArray[katPenzeszkozokArray.length - 1]) {
                        dataLength = jsonDataArray.length - 1;

                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["datum_meres_kezdete"] = item + "-01";
            postAsyncGetData(host + "/ebill/summary/getFeldolgozottMeresek", params, feldolgozottMeresekValuesCallback);
        };

        async.eachLimit(
            feldolgozottMeresekDateArray,
            threadLimit,
            feldolgozottMeresekValues,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getHHSzerzodes = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];
        // Ebben az értékben tároljuk a mértékegységet tartalmazó oszlopok értékeit
        var unitColumnValue;

        var HHSzerzodesCallback = function (err, HHSzerzodesCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
                setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
            }
            else {
                if (HHSzerzodesCallbackResult) {
                    var requiredServerDataArray = [
                        { dataTag: "id", columnName: "A", headerText: "ID" },
                        { dataTag: "elnevezes", columnName: "B", headerText: "Elnevezés" },
                        { dataTag: "ervenyesseg_kezdete", columnName: "C", headerText: "Szerz. kezdete" },
                        { dataTag: "ervenyesseg_vege", columnName: "D", headerText: "Szerz. vége" },
                        { dataTag: "halozati_engedelyes", columnName: "E", headerText: "Hálózati engedélyes" },
                        { dataTag: "meter_identifier_watt", columnName: "F", headerText: "Mérő azonosító" },
                        { dataTag: "POD", columnName: "G", headerText: "POD" },
                        { dataTag: "consumer_tariff_type", columnName: "H", headerText: "Tarifa" },
                        { dataTag: "lekotott_teljesitmeny", columnName: "I", headerText: "Lekötött teljesítmény" },
                        { dataTag: "lekotott_teljesitmeny_mertekegyseg", columnName: "J", headerText: "Mértékegység" },
                        { dataTag: "csatlakozasi_pontok_szama", columnName: "K", headerText: "Csatlakozási Pontok száma" },
                    ];

                    //Fejlécek betöltése a jsonDataArray-ba
                    jsonDataInnerArray = [];
                    jsonDataArray = [];
                    requiredServerDataArray.forEach(function (element) {
                        jsonDataInnerArray.push(element.headerText);
                    });
                    jsonDataArray.push(jsonDataInnerArray);
                    jsonDataInnerArray = [];

                    dataLength = Object.keys(HHSzerzodesCallbackResult.data).length;
                    dataInnerLength = requiredServerDataArray.length;

                    // Adattábla betöltése a jsonDataArray-ba
                    for (var tmpRow = 0; tmpRow < dataLength; tmpRow++) {
                        jsonDataInnerArray = [];
                        for (var i = 0; i < dataInnerLength; i++) {

                            switch (requiredServerDataArray[i].dataTag) {
                                case "id":
                                    HHHCSIds.push(HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                                    jsonDataInnerArray.push(HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                                    break;

                                case "csatlakozasi_pontok_szama":
                                    for (var j = 0; j < csatlakozasiPontResult.data.length; j++) {
                                        if (HHSzerzodesCallbackResult.data[tmpRow].id == csatlakozasiPontResult.data[j].id) {
                                            jsonDataInnerArray.push(csatlakozasiPontResult.data[j].csatlakozasi_pontok_szama);
                                            break;
                                        }
                                    }
                                    break;

                                case "lekotott_teljesitmeny":
                                    lekotottTeljesitmeny = HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag];
                                    if (lekotottTeljesitmeny == null) {
                                        indexOfSpace = -1;
                                    }
                                    else {
                                        indexOfSpace = lekotottTeljesitmeny.indexOf(" ");
                                    }
                                    if (indexOfSpace != -1) {
                                        jsonDataInnerArray.push(lekotottTeljesitmeny.substr(0, indexOfSpace))
                                        unitColumnValue = lekotottTeljesitmeny.substr(indexOfSpace + 1, lekotottTeljesitmeny.length)
                                    }
                                    else {
                                        jsonDataInnerArray.push("undefined");
                                        unitColumnValue = "undefined";
                                    }
                                    break;

                                case "lekotott_teljesitmeny_mertekegyseg":
                                    jsonDataInnerArray.push(unitColumnValue)
                                    unitColumnValue = "";
                                    break;

                                default:
                                    jsonDataInnerArray.push(HHSzerzodesCallbackResult.data[tmpRow][requiredServerDataArray[i].dataTag]);
                            }
                        }
                        jsonDataArray.push(jsonDataInnerArray);
                    }


                    excelDataArray.push(
                        {
                            "sheetName": "IN_VE1",
                            "data": jsonDataArray,
                        }
                    )
                    // ---------------------EXCEL RÉSZ ELEJE --------------------

                    //Excel.run(function (context) {

                    //    var sheet = context.workbook.worksheets.getItem("HHHCS szerződések");

                    //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                    //    //Excel feltöltése adatokkal
                    //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    //    range.values = jsonDataArray;
                    //    range.untrack();

                    //    // Csak a return után lesznek láthatóak az adatok az excelben
                    //    boldRange.format.font.bold = true;
                    //    return context.sync();
                    //})

                    // ---------------------EXCEL RÉSZ VÉGE --------------------
                    callback();

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
            }
        }

        var params = {};
        params["all"] = "1";
        params["isMasodlagos"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        postAsyncGetData(host + "/ebill/contract/getHHSzerzodes", params, HHSzerzodesCallback);

    }

    var getOperativTeljesitmeny = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var params = {};
        params["isMasodlagos"] = "0";
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "id", columnName: "A", headerText: "ID" }, // nem ennek a lekérdezésnek az ID-ja, hanem a hozzátartozó szerződésé
            { dataTag: "value", columnName: "B", headerText: "Érték" },
            { dataTag: "value_mertekegyseg", columnName: "C", headerText: "Mértékegység" },
            { dataTag: "ervenyesseg_kezdete", columnName: "D", headerText: "Érvényesség kezdete" },
            { dataTag: "ervenyesseg_vege", columnName: "E", headerText: "Érvényesség vége" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var operativTeljesitmenyek = function (item, innerCallback) {

            var operativTeljesitmenyekCallback = function (err, operativTeljesitmenyekCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = operativTeljesitmenyekCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                switch (requiredServerDataArray[j].dataTag) {
                                    case "id":
                                        jsonDataInnerArray.push(item);
                                        break;
                                    case "value":
                                        ertek = operativTeljesitmenyekCallbackResult[i][requiredServerDataArray[j].dataTag];
                                        if (ertek == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = ertek.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(ertek.substr(0, indexOfSpace))
                                            unitColumnValue = ertek.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "value_mertekegyseg":
                                        jsonDataInnerArray.push(unitColumnValue)
                                        unitColumnValue = "";
                                        break;

                                    default:
                                        jsonDataInnerArray.push(operativTeljesitmenyekCallbackResult[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray);

                        }
                    }
                    requestCounter++;

                    //if (item == HHHCSIds[HHHCSIds.length - 1]) {
                    if (requestCounter == HHHCSIds.length) {
                        dataLength = jsonDataArray.length - 1;

                        excelDataArray.push(
                            {
                                "sheetName": "IN_VE2",
                                "data": jsonDataArray,
                            }
                        )

                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("Operatív teljesítmény");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["operativ_szerzodes_id"] = item;
            postAsyncGetData(host + "/ebill/contract/Get_ebill_operativ_szerzodes", params, operativTeljesitmenyekCallback);
        };

        async.eachLimit(
            HHHCSIds,
            threadLimit,
            operativTeljesitmenyek,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getRendszerhasznalatiDijak = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;

        var rhdValuesArray = [];
        for (var i = 0; i < rhdValues.length; i++) {
            rhdValuesArray.push({ "kedo_datum": rhdValues[i][0].ervenyesseg_kezdete, "befejezo_datum": rhdValues[i][0].ervenyesseg_vege });
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "ervenyesseg_kezdete", columnName: "A", headerText: "Érvényesség kezdete" }, // Ennek az értéke nem a lekérdezésből jön
            { dataTag: "ervenyesseg_vege", columnName: "B", headerText: "Érvényesség vége" }, // Ennek az értéke nem a lekérdezésből jön
            { dataTag: "fogyaszto_tipus", columnName: "C", headerText: "Tarifa típus" },
            { dataTag: "elosztoi_alapdij", columnName: "D", headerText: "Elosztói alapdíj (Ft/csatl.p/év)" },
            { dataTag: "elosztoi_teljesitmeny_dij", columnName: "E", headerText: "Elosztói teljesítménydí (Ft/kW/év)" },
            { dataTag: "elosztoi_forgalmi_dij", columnName: "F", headerText: "Elosztói forgalmi díj (Ft/kWh)" },
            { dataTag: "elosztoi_meddo_energia_dij", columnName: "G", headerText: "Elosztói meddő energia díj (Ft/kVArh)" },
            { dataTag: "elosztoi_veszteseg_dij", columnName: "H", headerText: "Elosztói veszteség díj (Ft/kWh)" },
            { dataTag: "elosztoi_menetrend_kiegyensulyozasi_dij", columnName: "I", headerText: "Menetrend kiegyens. díj (Ft/kWh)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var rendszerHasznalatiDijak = function (item, innerCallback) {

            var rendszerHasznalatiDijakCallBack = function (err, rendszerHasznalatiDijakCallBackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = rendszerHasznalatiDijakCallBackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {

                                switch (requiredServerDataArray[j].dataTag) {
                                    case "ervenyesseg_kezdete":
                                        jsonDataInnerArray.push(item.kedo_datum);
                                        break;
                                    case "ervenyesseg_vege":
                                        jsonDataInnerArray.push(item.befejezo_datum);
                                        break;
                                    case "elosztoi_alapdij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }

                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_forgalmi_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_meddo_energia_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_menetrend_kiegyensulyozasi_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_teljesitmeny_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;
                                    case "elosztoi_veszteseg_dij":
                                        value = rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag];
                                        if (value == null) {
                                            indexOfSpace = -1;
                                        }
                                        else {
                                            indexOfSpace = value.indexOf(" ");
                                        }
                                        if (indexOfSpace != -1) {
                                            jsonDataInnerArray.push(value.substr(0, indexOfSpace))
                                            //unitColumnValue = penzeszkozEgysegar.substr(indexOfSpace + 1, ertek.length)
                                        }
                                        else {
                                            jsonDataInnerArray.push("undefined");
                                            //unitColumnValue = "undefined";
                                        }
                                        break;

                                    default:
                                        jsonDataInnerArray.push(rendszerHasznalatiDijakCallBackResult[i][requiredServerDataArray[j].dataTag]);
                                }
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }
                    requestCounter++;

                    //if (item.kedo_datum == rhdValuesArray[rhdValuesArray.length - 1].kedo_datum) {
                    if (requestCounter == rhdValuesArray.length) {
                        dataLength = jsonDataArray.length - 1;


                        excelDataArray.push(
                            {
                                "sheetName": "IN_VE3",
                                "data": jsonDataArray,
                            }
                        )
                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("RHD tarifafüggő");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item.kedo_datum;
            postAsyncGetData(host + "/ebill/admin/getRHDValues_mertekegyseggel", params, rendszerHasznalatiDijakCallBack);
        };

        async.eachLimit(
            rhdValuesArray,
            threadLimit,
            rendszerHasznalatiDijak,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    var getMeddoWattosValues = function (callback) {
        //Az excelbe bemásolandó range sorainak számát meghatározó változó
        var dataLength;
        //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
        var dataInnerLength;
        // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
        // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
        var jsonDataArray = [];
        // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
        var jsonDataInnerArray = [];

        threadLimit = 10;

        var innerCallbackDone = false;
        var requestCounter = 0;


        var meddoWattosArray = [];
        for (var i = 0; i < meddoWattosDatumok.length; i++) {
            meddoWattosArray.push(meddoWattosDatumok[i].ervenyesseg_kezdete);
        }

        var params = {};
        params["page"] = "1";
        params["start"] = "0";
        params["limit"] = "99999";

        var requiredServerDataArray = [
            { dataTag: "tarifatipus", columnName: "A", headerText: "Tarifa típus" },
            { dataTag: "ervenyesseg_kezdete", columnName: "B", headerText: "Érvényesség kezdete" },
            { dataTag: "arany", columnName: "C", headerText: "Arány (%)" },
        ];

        //Fejlécek betöltése a jsonDataArray-ba
        jsonDataInnerArray = [];
        jsonDataArray = [];
        requiredServerDataArray.forEach(function (element) {
            jsonDataInnerArray.push(element.headerText);
        });
        jsonDataArray.push(jsonDataInnerArray);
        jsonDataInnerArray = [];

        var meddoWattosValues = function (item, innerCallback) {

            var meddoWattosValuesCallback = function (err, meddoWattosValuesCallbackResult) {

                if (err) {
                    errorLabel.innerHTML = err.error.message;
                    changElementsAvailability(actualDisableElements, false);
                    setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                }
                else {
                    //Normálisan legenrálni a JSONArray változókat
                    dataLength = meddoWattosValuesCallbackResult.length;
                    dataInnerLength = requiredServerDataArray.length;
                    if (dataLength > 0) {
                        for (var i = 0; i < dataLength; i++) {
                            jsonDataInnerArray = [];
                            for (var j = 0; j < dataInnerLength; j++) {
                                jsonDataInnerArray.push(meddoWattosValuesCallbackResult[i][requiredServerDataArray[j].dataTag]);
                            }
                            jsonDataArray.push(jsonDataInnerArray);
                        }
                    }
                    requestCounter++;
                    //if (item == meddoWattosArray[meddoWattosArray.length - 1]) {
                    if (requestCounter == meddoWattosArray.length) {
                        dataLength = jsonDataArray.length - 1;

                        excelDataArray.push(
                            {
                                "sheetName": "IN_VE4",
                                "data": jsonDataArray,
                            }
                        )

                        // ---------------------EXCEL RÉSZ ELEJE --------------------

                        //Excel.run(function (context) {

                        //    var sheet = context.workbook.worksheets.getItem("Meddő-Wattos arány");

                        //    var boldRange = sheet.getRange("1:1").load("values, rowCount, columnCount");

                        //    //Excel feltöltése adatokkal
                        //    var range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                        //    range.values = jsonDataArray;
                        //    range.untrack();

                        //    // Csak a return után lesznek láthatóak az adatok az excelben
                        //    boldRange.format.font.bold = true;
                        //    return context.sync();
                        //})

                        // ---------------------EXCEL RÉSZ VÉGE --------------------


                        innerCallback();
                        innerCallbackDone = true;
                        callback();

                    }
                    if (innerCallbackDone == false) {
                        innerCallback();
                    }

                }
            }

            params["date_from"] = item;
            postAsyncGetData(host + "/ebill/admin/getMeddoWattosValues", params, meddoWattosValuesCallback);
        };

        async.eachLimit(
            meddoWattosArray,
            threadLimit,
            meddoWattosValues,
            function (err) {
                console.log('all finished', err);
            }
        );

    }

    //Munkafüzetet kezelő függvény

    var workSheetHandler = function (callback_lvl1) {

        var clearableSheet = [];
        var addableSheet = [];

        var separateWorksheets = function (callback_lvl2) {
            // Ez a függvény a lekérdezéshez szükséges munkalapokat két külön tömbe teszi.
            // A clearableSheet tömbbe teszi a már létező munkalapok nevét
            // Az addableSheet tömbbe teszi a létrehozandó munkalapok nevét
            Excel.run(function (context) {
                var worksheets = context.workbook.worksheets;
                worksheets.load('name');
                return context.sync()
                    .then(function () {
                        var sheetFound;
                        for (var i = 0; i < excelDataArray.length; i++) {
                            sheetFound = false;
                            for (var j = 0; j < worksheets.items.length; j++) {
                                if (excelDataArray[i].sheetName == worksheets.items[j].name) {
                                    sheetFound = true;
                                    clearableSheet.push(worksheets.items[j].name);
                                    break;
                                }
                            }
                            if (sheetFound) {
                                continue;
                            }
                            else {
                                addableSheet.push(excelDataArray[i].sheetName);
                            }

                        }
                        callback_lvl2();
                    });
            })

        }

        var clearSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var addSheets = function (callback_lvl2) {
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
                            callback_lvl2();
                        });
                });
            }
        }

        var loadDataToSheets = function (callback_lvl2) {
            if (excelDataArray) {
                Excel.run(function (context) {
                    var sheet;
                    var range;
                    var columnName;
                    var rowValue;

                    ////Munkalap nevének meghatározása
                    //sheet = context.workbook.worksheets.getItem("HHHCS szerződések");
                    ////Adatokkal feltöltendő tartomány meghatározésa
                    //range = sheet.getRange("A1:" + excelColumNames[dataInnerLength - 1] + (dataLength + 1));
                    ////Adatok feltöltése
                    //range.values = jsonDataArray;
                    //range.untrack();

                    for (var i = 0; i < excelDataArray.length; i++) {
                        sheet = context.workbook.worksheets.getItem(excelDataArray[i].sheetName);
                        columnName = excelColumNames[excelDataArray[i].data[0].length - 1];
                        rowValue = excelDataArray[i].data.length;

                        range = sheet.getRange("A1:" + columnName + rowValue);
                        range.values = excelDataArray[i].data;
                        range.untrack();

                    }


                    return context.sync()
                        .then(function () {
                            errorLabel.innerHTML = "";
                            errorLabel.style.display = "none";
                            changElementsAvailability(actualDisableElements, false);
                            setPanelLoader("rezsi-csokkentes-panel-loader", "rezsi-csokkentes-loader", "none");
                            callback_lvl2();
                            callback_lvl1();
                        });
                });
            }
        }

        async.series(
            [
                separateWorksheets,
                clearSheets,
                addSheets,
                loadDataToSheets
            ],
            function (err) {
                console.log('all finished', err);
            }
        );
    }

    var asyncSeriesFunctionsArray = [
        meterGroup,
        getCsatlakozasiPont,
        getMeddoWattosAranyValidFrom,
        getRHDValidFrom,
        getAllandoRendszerhasznalatiDijak,
        getFogyasztasOsszesito,
        getFeldolgozottMeresek,
        getHHSzerzodes,
        getOperativTeljesitmeny,
        getRendszerhasznalatiDijak,
        getMeddoWattosValues,
        workSheetHandler
    ];

    if (document.getElementById('rezsi_csokkentes_meter_groups').options.length == 0) {
        for (var i = 0; i < asyncSeriesFunctionsArray.length; i++) {
            if (asyncSeriesFunctionsArray[i] == getFogyasztasOsszesito) {
                asyncSeriesFunctionsArray.splice(i, 1);
                break;
            }
        }
    }

    async.series(
        asyncSeriesFunctionsArray,
        function (err) {
            console.log('allfinished', err);
        }
    )


}




