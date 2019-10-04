
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

//Excel cellák nevei egy tömbben (A,B,C,...)
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

// Azok az elementek (legfőképp a lenyíló menüket lenyitó gombok) amiket a lekérdezések alatt elérhetetlenné kell tenni, hogy a felhasználó ne tudja módosítani őket
var importantDisableElements = ["fogyasztasOsszesitoPanelOpen", "feldolgozottMeresekPanelOpen", "hetiJelentesPanelOpen"];

function fogyasztasOsszesitoContainer() {


    //Globális változó a meterGroup függvény eredményének kimentéséhez.
    var meterGroupArrayResult;
    //Lekérdezésekhez szükséges URL eleje
    var host = readCookie("enefexHost");

    var errorLabel = document.getElementById('fogyasztasOsszesitoError');
    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';


    // A függvényekben levő összes szükséges munkalapot itt kell definiálni
    var requiredSheets = ["IN_F0"];
    //Az excelbe bemásolandó range sorainak számát meghatározó változó
    var dataLength;
    //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
    var dataInnerLength;
    // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
    // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
    var jsonDataArray = [];
    // A jsonDataInnerArray tömb az amivel különböző ciklusokban feltöltjük a jsonDataArray változót
    var jsonDataInnerArray = [];

    // Lekérdezésekhez szükséges globális paraméterek eleje
    var dateFrom = document.getElementById('kezdo_datum').value;
    var dateTo = document.getElementById('veg_datum').value;

    // Lekérdezésekhez szükséges globális paraméterek vége

    //Dátumok RegEx validációi
    if (dateRegExTest('kezdo_datum', 'veg_datum', 'fogyasztasOsszesitoError') == "RegExTestProblem") {
        return;
    }

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    var newDisableElements = ["fogyasztasOsszesitoButton", "kezdo_datum", "veg_datum", "fogyasztas_osszesito_meter_groups"];

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
        params["limit"] = "25";

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

        var fogyasztasOsszesitoCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
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

                    // Menü elérhetővé tétele a lekérdezés végén
                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
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
        //Caolan async miatt
        callback();
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

    // A függvényekben levő összes szükséges munkalapot itt kell definiálni
    var requiredSheets = ["IN_É0"];
    //Az excelbe bemásolandó range sorainak számát meghatározó változó
    var dataLength;
    //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
    var dataInnerLength;
    // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
    // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
    var jsonDataArray = [];
    // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
    var jsonDataInnerArray = [];
    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");

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

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    var newDisableElements = ['onlyYearFilter', 'feldolgozott_meresek_meter_groups', 'feldolgozottMeresekButton'];
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
        params["limit"] = "25";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, getMeterGroupCallback);

    }

    var getFeldolgozottMeresek = function (callback) {

        var getFeldolgozottMeresekCallback = function (err, getFeldolgozottMeresekCallbackresult) {
            if (err) {
                errorLabel.innerHTML = err.error.message
                changElementsAvailability(actualDisableElements, false);
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

                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    errorLabel.innerHTML = "A getFeldolgozottMeresekCallback resultjában lévő JSON Object hibás vagy üres"
                    changElementsAvailability(actualDisableElements, false);
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
        params["limit"] = "1000";

        postAsyncGetData(host + "/ebill/summary/getFeldolgozottMeresek", params, getFeldolgozottMeresekCallback);

        //Caolan async miatt
        callback();
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

    // A függvényekben levő összes szükséges munkalapot itt kell definiálni
    var requiredSheets = ["IN_FÖ", "IN_SzA", "IN_FG"];
    //Az excelbe bemásolandó range sorainak számát meghatározó változó
    var dataLength;
    //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
    var dataInnerLength;
    // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
    // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
    var jsonDataArray = [];
    // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
    var jsonDataInnerArray = [];
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
        requiredSheets.splice(2, 1);
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
        return;
    }

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    var newDisableElements = ['heti_jelentes_kezdo_datum', 'heti_jelentes_kezdo_ora', 'heti_jelentes_veg_datum', 'heti_jelentes_befejezo_ora', 'heti_jelentes_meter_groups', 'heti_jelentes_mentett_bealitasok', 'csakNemMegFeleloSorok', 'hetiJelentesButton'];

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
        params["limit"] = "25";

        postAsyncGetData(host + "/ebill/billing/getMeterGroups", params, getMeterGroupCallback);

    }

    var getMeterTree = function (callback) {

        var metreGroupCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
            }
            else {
                if (result) {
                    meterTreeArray = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
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
            }
            else {
                if (result) {
                    savedOptionsArray = result;
                    callback();
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás" 
                    changElementsAvailability(actualDisableElements, false);
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

        var hetiFogyasztasOsszesitoCallback = function (err, hetiFogyasztasOsszesitoCallbackResult) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
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

                    //Menü elérhetővé tétele
                    changElementsAvailability(actualDisableElements, false);
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
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
        //Caolan async miatt
        callback();
    }

    var getHetiAlapvonalSzamertekek = function (callback) {
        var hetiAlapvonalSzamertekekCallback = function(err, result){
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
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
                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás"
                    changElementsAvailability(actualDisableElements, false);
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
        params["limit"] = "25";
        params["sort"] = '[{ "property": "T.id", "direction": "desc" }]';

        postAsyncGetData(host + "/vstat/baseline/getTableStat", params, hetiAlapvonalSzamertekekCallback);

        //Caolan async miatt
        callback();
    }

    var getMentettBeallitasokGrafikonAdatok = function (callback) {
        var mentettBeallitasokGrafikonAdatokCallback = function (err, result) {
            if (err) {
                errorLabel.innerHTML = err.error.message;
                changElementsAvailability(actualDisableElements, false);
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

                }
                else {
                    errorLabel.innerHTML = "A szerverről lekért JSON Object üres vagy hibás";
                    changElementsAvailability(actualDisableElements, false);
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
        //Caolan async miatt
        callback();
    }

    var asyncSeriesFunctionsArray = [workSheetHandler, meterGroup, getMeterTree, getSavedGraphs, getHetiAlapvonalSzamertekek, getMentettBeallitasokGrafikonAdatok, getHetiFogyasztasOsszesito];

    if (!savedOptionsListSelectedText) {
        asyncSeriesFunctionsArray.splice(5, 1);
    }

    async.series(
        //[
        //    // Adatokat tartalmazó JSON lekérdezések paramétereit meghatározó függvények
        //    meterGroup,
        //    getMeterTree,
        //    getSavedGraphs,
        //    workSheetHandler,
        //    getHetiAlapvonalSzamertekek,
        //    getMentettBeallitasokGrafikonAdatok,
        //    //Mindenképp a getHetiFogyasztasOsszesito legyen az utolsó async lekérdezés
        //    getHetiFogyasztasOsszesito
        //],
        asyncSeriesFunctionsArray,
        function (err) {
            console.log('allfinished', err);
        }
    )

}

//MINTA ASYNC függvény

function mintaAsync() {



    var errorLabel = document.getElementById('AKARMIERROR');
    errorLabel.style.display = "block";
    errorLabel.innerHTML = '<span class="green-text">Szerverlekérdezés folymatban...</span>';

    // A függvényekben levő összes szükséges munkalapot itt kell definiálni
    var requiredSheets = ["TEST2", "TEST1"];
    //Az excelbe bemásolandó range sorainak számát meghatározó változó
    var dataLength;
    //Az excelbe bemásolandó range oszlopainak számát meghatározó változó
    var dataInnerLength;
    // Az excelbe való adatbevitelkor ha rangebe akarjuk megadni a beírandó adatokat akkor azt egy több dimenziójú tömb változóba tehetjük
    // A jsonDataArray tömb az amit az excelnek megadunk mint beírandó tömb
    var jsonDataArray = [];
    // A  jsonDataInnerArray tömb az amivel ciklusonként feltöltjük a jsonDataArray változót
    var jsonDataInnerArray = [];
    // Lekérdezésekhez szükséges URL
    var host = readCookie("enefexHost");

    //!!!!!!!!!!!!!!!!!!!!!!
    // Egyéb szükséges paraméterek definiálása (input typeok tartalmai...)

    //!!!!!!!!!!!!!!!!!!!!!!

    //Dátumok RegEx validációi
    if (dateRegExTest('kezdo_datum_input', 'veg_datum_input', 'AKARMIERROR') == "RegExTestProblem") {
        return;
    }

    //Menü elérhetetlenné tétele a lekérdezés alatt, hogy a felhasználó ne tudja elcseszni
    var newDisableElements = ["disablable_input", "disablable_Button"];
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

    // Többi Async függvény definiálása

    //Először azokat a függvényeket kell megírni amikből a főbb függvények lekérdezéséhez szükséges paramértert lehet kinyerni

    // A megírt lekérdezések közül csak az UTOLSÓBAN az EXCELBEÍRÁS UTÁN (és minden errorágon belül) kell aktiválni a kezelőfelületi elmeket.
    // Kezelőfelületeket aktiváló kód: changElementsAvailability(actualDisableElements, false);



    async.series(
        [
            workSheetHandler,
            // többi async függvény
        ],
        function (err) {
            console.log('allfinished', err);
        }
    )

}