(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;
    var eventResult;
    
    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight 2!");
            $('#button-desc').text("Highlights the largest number.");
                
            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(insertImageBase64FromRange);

            $('#table-button-text').text("insert table");
            $('#table-button-desc').text("table description");

            // Add a click event handler for the highlight button.
            $('#table-button').click(insertTableToSheet);

            //Excel.run(function (context) {
            //    var worksheet = context.workbook.worksheets.getActiveWorksheet();
            //    eventResult = worksheet.onSelectionChanged.add(handleSelectionChange);

            //    return context.sync()
            //        .then(function () {
            //            console.log("Event handler successfully registered for onSelectionChanged event in the worksheet.");
            //        });
            //}).catch(errorHandler);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;
            //sheet.insertFileFromBase64();

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function handleSelectionChange(event) {
        return Excel.run(function (context) {
            return context.sync()
                .then(function () {
                    console.log("Address of current selection: " + event.address);
                });
        }).catch(errorHandler);
    }

    function ToggleEnabledEvents() {
        Excel.run(function (context) {
            context.runtime.load("enableEvents");
            return context.sync()
                .then(function () {
                    var eventBoolean = !context.runtime.enableEvents;
                    context.runtime.enableEvents = eventBoolean;
                    if (eventBoolean) {
                        console.log("Events are currently on.");
                    } else {
                        console.log("Events are currently off.");
                    }
                }).then(context.sync);
        }).catch(errorHandler);
    }


    function removeEvent() {
        return Excel.run(eventResult.context, function (context) {
            eventResult.remove();

            return context.sync()
                .then(function () {
                    eventResult = null;
                    console.log("Event handler successfully removed.");
                });
        }).catch(errorHandler);
    }

    function insertTableToSheet() {
        Excel.run(function (ctx) {

            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var table = ctx.workbook.tables.add(sheet.getRange("A1:E7"), true);

            table.load('name', 'columns', 'rows');

            return ctx.sync().then(function () {

                var col = table.columns.getItemAt(0).values;
                col.load('values');
                
                return ctx.sync().then(function () {

                    console.log(col.values);
                    });
                });
        }).catch(errorHandler);
    }


    
    function hightlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight 
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    function insertImageBase64FromRange() {
        Excel.run(function (ctx) {

            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            var address = "B3:D5";
            var range = sheet.getRange(address);
            var base64String = range.getImage();

            // Må kjørr sync for å hente clientresult.value fra getimage()
            return ctx.sync().then(function () {
                console.log(base64String.value);
                sheet.shapes.addImage(base64String.value);

            }).then(ctx.sync);

        }).catch(errorHandler);
    }


    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
