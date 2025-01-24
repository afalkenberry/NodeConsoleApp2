'use strict';

const headers = new Headers({
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*', // Allow all origins for testing
});

const requestOptions = {
    method: 'GET',
    headers: headers,
};

fetch('http://localhost:41951/DYMO/DLS/Printing/StatusConnected', requestOptions)
    .then(response => response.json())
    .then(data => console.log('Service is running:', data))
    .catch(error => console.error('Error:', error));

window.addEventListener('DOMContentLoaded', function () {
    if (typeof dymo === 'undefined' || !dymo.label || !dymo.label.framework) {
        console.error('DYMO SDK not loaded!');
        return;
    }

    console.log('DYMO SDK is loaded successfully!');
});

// declare cached global variables
let cachedCompleteQRString = null;
let cachedPrinterName = null;
let cachedPartNumber = null;
let cachedSNQty = null;

//Upload excel file
document.getElementById('upload-button').addEventListener('click', function () {
    document.getElementById('my_file').click();
});

// Process the .xlsx file immediately after it's selected
document.getElementById('my_file').addEventListener('change', function () {
    const fileInput = this;
    const output = document.getElementById('output');
    const qrOutput = document.getElementById('outputQRString');

    // Check if a file was selected
    if (fileInput.files.length === 0) {
        output.textContent = 'No file selected.';
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();

    // Read the file as binary string
    reader.onload = function (event) {
        try {
            const data = event.target.result;

            // Parse the excel
            const workbook = XLSX.read(data, { type: 'binary' });

            // Get the excel sheet
            const sheetName = workbook.SheetNames[0];
            const sheet = workbook.Sheets[sheetName];

            // Convert excel sheet data to JSON
            const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

            // Flatten the data into a 1D array
            const flattenedData = jsonData.flat().filter(item => item != null && item.toString().trim() !== "");

            //console.log('Flattened Array:', flattenedData);

            //Put SNs into format needed for QR string
            const header = jsonData[0]?.[0];
            //console.log('Header:', header);

            //Create header for QR string
            const snArray = jsonData.flat().slice(1).filter(item => item != null && item.toString().trim() !== "");
            const count = snArray.length;
            const qrHeader = ("[)>" + "{RS}" + count + "{GS}" + header + "{GS}");

            // store qty of SN and part number in cached variables
            cachedSNQty = count;
            cachedPartNumber = header;

            //Add {GS} to serial numbers in array (body of QR string)
            function addGSAndJoin(array) {
                // Get the first item
                const firstItem = array[0];

                // Skip the first item, append {GS} to the remaining items, then join them into a string
                const modifiedString = array
                    .slice(1) // Skip the first item
                    .map(item => `${item}{GS}`) // Add {GS} to each item
                    .join(''); // Join them into a single string

                // Add the first item to the beginning of the string
                return `${firstItem}{GS}${modifiedString}`;
            }

            // store SNs with {GS} appended in variable
            const qrSNS = addGSAndJoin(snArray);

            // Add everything together to make the complete QR String format
            cachedCompleteQRString = (qrHeader + qrSNS + "{RS}" + "{EoT}");

            // Output the flattened array to the screen
            output.textContent = `Excel Input:${JSON.stringify(flattenedData, null, 2)}`;
        } catch (e) {
            console.error('Error processing file:', e);
            output.textContent = 'Error processing file.';
        }
    };

    reader.onerror = function () {
        console.error('Error reading file');
        output.textContent = 'Error reading file.';
    };

    reader.readAsBinaryString(file);
});

function preloadPrinters() {
    const printers = dymo.label.framework.getPrinters();
    if (printers.length > 0) {
        cachedPrinterName = printers[0].name; // Choose the first printer
    } else {
        alert("No DYMO printers are installed or connected.");
    }
}
function checkPrinterAvailability(printerName) {
    if (!printerName) {
        alert("No printer available. Please check your setup.");
        return;
    }
}

// Call preloadPrinters to set cachedPrinterName
preloadPrinters();

// Check if the printer is available
checkPrinterAvailability(cachedPrinterName);

// Label xml for qr code
const baseLabelXml = '<?xml version="1.0" encoding="utf-8"?> \
                    <DieCutLabel Version="8.0" Units="twips"> \
	                    <PaperOrientation>Landscape</PaperOrientation> \
	                    <Id>NameBadgeTag</Id> \
	                    <PaperName>30252 Address</PaperName> \
	                    <ObjectInfo>  \
                            <BarcodeObject >  \
                                <Name>Barcode</Name> \
                                <ForeColor Alpha="255" Red="0" Green="0" Blue="0" /> \
                                <BackColor Alpha="0" Red="255" Green="255" Blue="255" /> \
                                <LinkedObjectName>BarcodeText</LinkedObjectName> \
                                <Rotation>Rotation0</Rotation> \
                                <IsMirrored>False</IsMirrored> \
                                <IsVariable>True</IsVariable> \
                                <Text>{DYNAMIC_TEXT}</Text> \
                                <Type>QRCode</Type> \
                                <Size>Small</Size> \
                                <TextPosition>Bottom</TextPosition> \
                                <TextFont Family="Arial" Size="8" Bold="False" Italic="False" Underline="False" Strikeout="False" /> \
                                <CheckSumFont Family="Arial" Size="8" Bold="False" Italic="False" Underline="False" Strikeout="False" /> \
                                <TextEmbedding>None</TextEmbedding> \
                                <ECLevel>0</ECLevel> \
                                <HorizontalAlignment>Left</HorizontalAlignment> \
                                <QuietZonesPadding Left="0" Top="0" Right="0" Bottom="0" /> \
                            </BarcodeObject > \
                            <Bounds X="0" Y="0" Width="5040.32258" Height="1620.10369" /> \
                        </ObjectInfo > \
                        <ObjectInfo> \
                            <TextObject> \
                                <Name>BarcodeText</Name> \
                                <ForeColor Alpha="255" Red="0" Green="0" Blue="0" /> \
                                <BackColor Alpha="0" Red="255" Green="255" Blue="255" /> \
                                <LinkedObjectName></LinkedObjectName> \
                                <Rotation>Rotation0</Rotation> \
                                <IsMirrored>False</IsMirrored> \
                                <IsVariable>True</IsVariable> \
                                <HorizontalAlignment>Right</HorizontalAlignment> \
                                <VerticalAlignment>Top</VerticalAlignment> \
                                <TextFitMode>ShrinkToFit</TextFitMode> \
                                <UseFullFontHeight>False</UseFullFontHeight> \
                                <Verticalized>False</Verticalized> \
                                <StyledText> \
                                    <Element> \
                                        <String>{DYNAMIC_TEXT2}</String> \
                                        <Attributes> \
                                            <Font Family="Arial" Size="7" Bold="False" Italic="False" Underline="False" Strikeout="False" /> \
                                            <ForeColor Alpha="255" Red="0" Green="0" Blue="0" /> \
                                        </Attributes> \
                                    </Element> \
                                </StyledText> \
                            </TextObject> \
                            <Bounds X="1800" Y="0" Width="5040.32258" Height="1620.10369" /> \
                        </ObjectInfo> \
                        <ObjectInfo> \
                            <TextObject> \
                                <Name>BarcodeText</Name> \
                                <ForeColor Alpha="255" Red="0" Green="0" Blue="0" /> \
                                <BackColor Alpha="0" Red="255" Green="255" Blue="255" /> \
                                <LinkedObjectName></LinkedObjectName> \
                                <Rotation>Rotation0</Rotation> \
                                <IsMirrored>False</IsMirrored> \
                                <IsVariable>True</IsVariable> \
                                <HorizontalAlignment>Right</HorizontalAlignment> \
                                <VerticalAlignment>Bottom</VerticalAlignment> \
                                <TextFitMode>ShrinkToFit</TextFitMode> \
                                <UseFullFontHeight>False</UseFullFontHeight> \
                                <Verticalized>False</Verticalized> \
                                <StyledText> \
                                    <Element> \
                                        <String>Qty: {DYNAMIC_TEXT3}</String> \
                                        <Attributes> \
                                            <Font Family="Arial" Size="7" Bold="False" Italic="False" Underline="False" Strikeout="False" /> \
                                            <ForeColor Alpha="255" Red="0" Green="0" Blue="0" /> \
                                        </Attributes> \
                                    </Element> \
                                </StyledText> \
                            </TextObject> \
                            <Bounds X="3360.21505" Y="0" Width="5040.32258" Height="1620.10369" /> \
                        </ObjectInfo> \
                        </DieCutLabel>';

function printLabel() {

    if (!cachedPrinterName || !cachedCompleteQRString) {
        console.error("Printer or QR String not initialized.");
        return;
    }

    try {
        const labelXml = baseLabelXml.replace("{DYNAMIC_TEXT}", cachedCompleteQRString).replace("{DYNAMIC_TEXT2}", cachedPartNumber).replace("{DYNAMIC_TEXT3}", cachedSNQty);
        const label = dymo.label.framework.openLabelXml(labelXml);
        label.print(cachedPrinterName);
        alert("Label sent to printer!");

    } catch (error) {
        console.error("Error printing label:", error);
    }
}

function getPrintersWithHeaders() {
    fetch('http://localhost:41951/DYMO/DLS/Printing/GetPrinters', {
        method: 'GET',
        headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*', // Include appropriate headers
        },
    })
        .then(response => response.json())
        .then(printers => {
            console.log('Printers:', printers);
        })
        .catch(error => console.error('Error fetching printers:', error));
}

// Call this function instead of dymo.label.framework.getPrinters
getPrintersWithHeaders();
