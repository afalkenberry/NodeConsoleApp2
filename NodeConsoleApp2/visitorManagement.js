
(function () {

    // return list of available layouts, load them from file if nessesary
    function getLayouts() {
        if (_layouts == null) {
            // load layouts
            var result = new Array();
            var base = document.location.href;
            for (var i = 0; i < _layoutFiles.length; i++)
                result.push(dymo.label.framework.openLabelXml(_layoutFiles[i]));
            _layouts = result;
        }

        return _layouts;
    }

    // current data on the label. For simplicity we support one Text obje t and one Image object
    var _labelData = { text: "Name and other information" };

    // applies data to the label
    function applyDataToLabel(label, labelData) {
        var names = label.getObjectNames();

        for (var name in labelData)
            if (itemIndexOf(names, name) >= 0)
                label.setObjectText(name, labelData[name]);
    }

   

    // returns an index of an item in an array. Returns -1 if not found
    function itemIndexOf(array, item) {
        for (var i = 0; i < array.length; i++)
            if (array[i] == item) return i;

        return -1;
    }

    // loads the defualt layout at onload()
    function setupDefaultLayout() {
        _label = dymo.label.framework.openLabelXml(_layoutFiles[0]);
        applyDataToLabel(_label, _labelData);
    }

    // called when clicked on a photo
    
    function createPrintersTableRow(table, name, value) {
        var row = document.createElement("tr");

        var cell1 = document.createElement("td");
        cell1.appendChild(document.createTextNode(name + ': '));

        var cell2 = document.createElement("td");
        cell2.appendChild(document.createTextNode(value));

        row.appendChild(cell1);
        row.appendChild(cell2);

        table.appendChild(row);
    }

    function populatePrinterDetail() {
        var printerDetail = document.getElementById("printerDetail");
        printerDetail.innerHTML = "";

        var myPrinter = _printers[document.getElementById("printersSelect").value];
        if (myPrinter === undefined)
            return;

        var table = document.createElement("table");
        createPrintersTableRow(table, 'PrinterType', myPrinter['printerType'])
        createPrintersTableRow(table, 'PrinterName', myPrinter['name'])
        createPrintersTableRow(table, 'ModelName', myPrinter['modelName'])
        createPrintersTableRow(table, 'IsLocal', myPrinter['isLocal'])
        createPrintersTableRow(table, 'IsConnected', myPrinter['isConnected'])
        createPrintersTableRow(table, 'IsTwinTurbo', myPrinter['isTwinTurbo'])

        dymo.label.framework.is550PrinterAsync(myPrinter.name).then(function (isRollStatusSupported) {
            //fetch one consumable information in the printer list.
            if (isRollStatusSupported) {
                createPrintersTableRow(table, 'IsRollStatusSupported', 'True')
                dymo.label.framework.getConsumableInfoIn550PrinterAsync(myPrinter.name).then(function (consumableInfo) {
                    createPrintersTableRow(table, 'SKU', consumableInfo['sku'])
                    createPrintersTableRow(table, 'Consumable Name', consumableInfo['name'])
                    createPrintersTableRow(table, 'Labels Remaining', consumableInfo['labelsRemaining'])
                    createPrintersTableRow(table, 'Roll Status', consumableInfo['rollStatus'])
                }).thenCatch(function (e) {
                    createPrintersTableRow(table, 'SKU', 'n/a')
                    createPrintersTableRow(table, 'Consumable Name', 'n/a')
                    createPrintersTableRow(table, 'Labels Remaining', 'n/a')
                    createPrintersTableRow(table, 'Roll Status', 'n/a')
                })
            } else {
                createPrintersTableRow(table, 'IsRollStatusSupported', 'False')
            }
        }).thenCatch(function (e) {
            createPrintersTableRow(table, 'IsRollStatusSupported', e.message)
        })

        printerDetail.appendChild(table);
    }

    // called when the document completly loaded
    function onload() {
        var labelFile = document.getElementById('labelFile');
        var labelTextTextArea = document.getElementById('labelTextTextArea');
        var printersSelect = document.getElementById('printersSelect');
        var printButton = document.getElementById('printButton');
        var selectPhotoButton = document.getElementById('selectPhotoButton');
        var changeLayoutButton = document.getElementById('changeLayoutButton');


        // initialize controls
        //printButton.disabled = true;
        //addressTextArea.disabled = true;
        if (_labelData.text)
            labelTextTextArea.value = _labelData.text;

        // loads all supported printers into a combo box 
        function loadPrintersAsync() {
            _printers = [];
            dymo.label.framework.getPrintersAsync().then(function (printers) {
                if (printers.length == 0) {
                    alert("No DYMO printers are installed. Install DYMO printers.");
                    return;
                }
                _printers = printers;
                printers.forEach(function (printer) {
                    let printerName = printer["name"];
                    let option = document.createElement("option");
                    option.value = printerName;
                    option.appendChild(document.createTextNode(printerName));
                    printersSelect.appendChild(option);
                });
                populatePrinterDetail();
            }).thenCatch(function (e) {
                alert("Load Printers failed: " + e);;
                return;
            });
        };

        // updates address on the label when user types in textarea field
        labelTextTextArea.onkeyup = function () {
            if (!_label) {
                alert('Load label before entering text');
                return;
            }

            // set labelData
            _labelData.text = labelTextTextArea.value;
            applyDataToLabel(_label, _labelData);
            updatePreview();
        }

        // prints the label
        printButton.onclick = function () {
            if (!_label) {
                alert("Load label before printing");
                return;
            }

            //alert(printersSelect.value);
            _label.print(printersSelect.value);
        }

        selectPhotoButton.onclick = selectPhotoButtonClick;
        changeLayoutButton.onclick = changeLayoutButtonClick;
        printersSelect.onchange = populatePrinterDetail;

        // onload() initialization
        loadPrintersAsync();
        setupDefaultLayout();
        updatePreview();
        updateControls();
    };

    function initTests() {
        if (dymo.label.framework.init) {
            //dymo.label.framework.trace = true;
            dymo.label.framework.init(onload);
        } else {
            onload();
        }
    }

    // register onload event
    if (window.addEventListener)
        window.addEventListener("load", initTests, false);
    else if (window.attachEvent)
        window.attachEvent("onload", initTests);
    else
        window.onload = initTests;

}());