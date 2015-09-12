/// <reference path="../App.js" />
// global app
//  Published at https://azureiotservice.azurewebsites.net/UpdateSensorData.zip

//  The next row number to be populated.
var nextRow = 0;

//  The values to be populated in the new row, e.g. ["2015-09-11 16:23:42", 28.5, 185]
var newRow = null;

(function () {
    'use strict';

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
			
			//  TODO: When pause button is clicked, pause the updating.
			$('#pause').click(createTable);
			//  TODO: When resume button is clicked, resume the updating.
			//$('#resume').click(AddRowsToTable);
			
			//  Create the table if it doesn't exist.
			createTable();
			
			//  Get a new row from AzureIoTService every 10 seconds.
			setInterval(getNewRow, 10000);

        });
    };
	
	function createTable() { 
		//  Create the table if it doesn't exist.
		Office.context.document.bindings.getByIdAsync("myTable", function (asyncResult) { 
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { 
				return;
            }
			var myTable = new Office.TableData(); 
	        myTable.headers = [["Timestamp", "Temperature", "LightLevel"]]; 
	        //myTable.rows = null; 
			nextRow = 2;		
	        Office.context.document.setSelectedDataAsync(myTable, 
	            { tableOptions: { bandedRows: true, filterButton: true, style: "TableStyleMedium2" } }, 
	            bindTable);
        });
    }

    function bindTable() { 
        Office.context.document.bindings.addFromSelectionAsync(Office.BindingType.Table, { id: 'myTable' }, function (asyncResult) { 
            if (asyncResult.status === Office.AsyncResultStatus.Failed) { 
                var error = asyncResult.error; 
                app.showNotification("Error", error.name + ": " + error.message); 
            } else { 
                app.showNotification("OK"); 
            } 
        }); 
    } 
 
    // Add rows to an existing table 
    function newRowLoaded() { 
		//  Load the new row.
		app.showNotification(newRow);
		//  If no new row to add, try again later.
		if (!newRow) return;
         
        Office.context.document.bindings.getByIdAsync("myTable", function (asyncResult) { 
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) { 
                var currentTable = asyncResult.value; 
                var rowCount = currentTable.rowCount; 
				if (nextRow == 2) {
					replaceRow();
					return;
				}
                //app.showNotification(rowCount); 
				
				//  Add the new row, which is something like ["2015-09-11T16:23:42", 28.5, 185]
				currentTable.addRowsAsync([newRow], 
                    function (e) { 
						var json = JSON.stringify(newRow);
                        if (e.status === Office.AsyncResultStatus.Failed) { 
                            app.showNotification(json + " / Error", e.error.name + ": " + e.error.message); 
                        } else { 
					        app.showNotification(json + ' / OK');
                        } 
                 });
            } 
        }); 
    } 

	function displayAllBindingNames() {
	    Office.context.document.bindings.getAllAsync(function (asyncResult) {
	        var bindingString = '';
			var count = 0;
	        for (var i in asyncResult.value) {
	            bindingString += asyncResult.value[i].id + ' / \n';
				count++;
	        }
			app.showNotification(count + " bindings: \n" + bindingString);
	    });
	}
	
	function getNewRow() {
		//  Insert dynamic JavaScript to get new row values from server.  Based on http://www.hunlock.com/blogs/Howto_Dynamically_Insert_Javascript_And_CSS 
		var headID = document.getElementsByTagName("head")[0];         
		var newScript = document.createElement('script');
		newScript.type = 'text/javascript';
		//  Wait for script to execute, then load the new row.
		newScript.onload = newRowLoaded;
		//  The script will return something like
		//  newRow = ["2015-09-11T16:23:42", 28.5, 185];
		//  Must use HTTPS because the calling HTML page (Excel Online) is also HTTPS.
		//  TODO: Send the group name and column headers to the service.
		newScript.src = 'https://azureiotservice.azurewebsites.net/Test.aspx';
		headID.appendChild(newScript);		
	}
	
	function replaceRow() {
		//  Replace the second row by the new row data.
		app.showNotification(newRow);
		//  If no new row to add, try again later.
		if (!newRow) return;
		var prevNextRow = nextRow;
		nextRow++;
		//  Bind to a new row in the table, e.g. A1:C1, A2:C2, A3:C3, ...
        Office.context.document.bindings.addFromNamedItemAsync("Sheet1!A" + prevNextRow + ":C" + prevNextRow, "matrix", {id: "ReplaceRowBinding" },
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification('Error: ' + asyncResult.error.message);
            }
            else {
                //  Populate the bound row.
				//  newRow is something like ["2015-09-11T16:23:42", 28.5, 185]
                Office.select("bindings#ReplaceRowBinding").setDataAsync([newRow], { coercionType: "matrix" },
                    function (asyncResult) {
						var json = JSON.stringify(newRow);
                        if (asyncResult.status == "failed") {
                            app.showNotification(json + ' / Error: ' + asyncResult.error.message);
                        } else {
					        app.showNotification(json + ' / OK');
						}
                    });
            }
        });
	}
	
})();

	/*
    // we use a queue for applying the cell formatting because of Excel online limitations 
    // which restrict cell formatting calls to 100 cells or less 
    rangeFormatQueue = []; 
    rangeFormatQueue.push({ cells: Office.Table.All, format: { width: "auto fit" } }); 
    for (var j = 0; j < myTable.rows.length; j++) { 
        rangeFormatQueue.push({ cells: { row: j, column: 1 }, format: { numberFormat: dateFormat } }); 
        rangeFormatQueue.push({ cells: { row: j, column: 3 }, format: { fontColor: myTable.rows[j][3] } }); 
    } 
	*/

	/*
    rangeFormatQueue = []; 
    rangeFormatQueue.push({ cells: Office.Table.All, format: { width: "auto fit" } }); 
    for (var j = 0; j < rows.length; j++) { 
        rangeFormatQueue.push({ cells: { row: j + rowOffset, column: 1 }, format: { numberFormat: "dd-mmm-yyyy" } }); 
        rangeFormatQueue.push({ cells: { row: j + rowOffset, column: 3 }, format: { fontColor: rows[j][3] } }); 
    } 
	*/

	/*
    // Reads data from current document selection and displays a notification
    function NOTUSED_getDataFromSelection() {
		//Write text to the current user selection
		Office.context.document.setSelectedDataAsync(
		  nextRow++, 
		  function (asyncResult) {
		    if (asyncResult.status == "failed") {
		      app.showNotification("Action failed with error: " + asyncResult.error.message);
		    } else {
		      app.showNotification("Success! Click the Next button to move on.");
		  }
		});
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
			function (result) {
			    if (result.status === Office.AsyncResultStatus.Succeeded) {
			        app.showNotification('The selected text is:', '"' + result.value + '"');
			    } else {
			        app.showNotification('Error:', result.error.message);
			    }
			}
		);
    }
	*/
