function createImportSheets() {
	var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
	var importSheets=["main","major","degree","interest","test"]
	importSheets.reverse() // Creates the spreadsheets in the desired order
	for (var i = importSheets.length - 1; i >= 0; i--) {  // For each sheet to create
		sheet = spreadsheet.getSheetByName(importSheets[i]);
		// Only create it if it doesn't exist
		if (sheet == null) { 
			spreadsheet.insertSheet(importSheets[i], spreadsheet.getSheets().length);
		}
	}
	SpreadsheetApp.getActiveSpreadsheet().getSheetByName("all").activate(); // Go back to the all tab
}

function createInterestSheets(){
	var id  = SpreadsheetApp.getActiveSpreadsheet().getId();// get the actual id
	// Get unique list of interests
	var interest_sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("interest");
	var end_data_row=interest_sheet.getLastRow()
	// Select all interests cells
	var interests=interest_sheet.getRange("B2:B"+end_data_row).getValues();
	// Convert to flattened array
	var interests_flat=[].concat.apply([], interests);
	var unique_interests=[]
	// For each cell if the interest isn't in the unique_interests array add it
	for (var row = 0; row < interests.length; row++) {
		if (interests_flat[row] != '' && unique_interests.indexOf(interests_flat[row]) == -1) {
			unique_interests.push(interests_flat[row]);
		}
	}
	var long_titles=SpreadsheetApp.openById(id).getSheetByName("lookups").getRange("E2:E").getValues();
	var long_titles=[].concat.apply([], long_titles);
	// Get Short Titles
	var short_titles=SpreadsheetApp.openById(id).getSheetByName("lookups").getRange("F2:F").getValues();
	var short_titles=[].concat.apply([], short_titles);
	// Create sheets if they don't exist
	var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
	for (var i = unique_interests.length - 1; i >= 0; i--) {  // For each sheet to create
		// If we have a short name use it
		if(long_titles.indexOf(unique_interests[i])>-1){
			var short_name=short_titles[long_titles.indexOf(unique_interests[i])]
			Logger.log(short_name)
			sheet = spreadsheet.getSheetByName(short_name);
			// Only create it if it doesn't exist
			if (sheet == null) { 
				spreadsheet.insertSheet(short_name, spreadsheet.getSheets().length);
				sheet = spreadsheet.getSheetByName(short_name);
			}
		}
		// Else
		else{
			sheet = spreadsheet.getSheetByName(unique_interests[i]);
			// Only create it if it doesn't exist
			if (sheet == null) { 
				spreadsheet.insertSheet(unique_interests[i], spreadsheet.getSheets().length);
				sheet = spreadsheet.getSheetByName(unique_interests[i]);
			}
		}
		sheet.getRange('A2').setValue("=ARRAYFORMULA(all!2:2)")
	}
	SpreadsheetApp.getActiveSpreadsheet().getSheetByName("all").activate(); // Go back to the all tab
}

function processData(){
	var id  = SpreadsheetApp.getActiveSpreadsheet().getId();// get the actual id
	var importSheets=["main","degree","interest","test"]
	var all=SpreadsheetApp.openById(id).getSheetByName("all");
	var main=SpreadsheetApp.openById(id).getSheetByName("main");
	var interest=SpreadsheetApp.openById(id).getSheetByName("interest");
	var start_data_row=3
	all.getRange('L1').setValue("1. Converting MyGrad text format data to spreadsheet format")
	for (var i = importSheets.length - 1; i >= 0; i--) {  // For each admiss spreadsheet convert admissions style import to usable spreadsheet
		var spreadsheet = SpreadsheetApp.openById(id).getSheetByName(importSheets[i])
		var numRows = spreadsheet.getLastRow()  // Only do it for as many rows as are needed
		var formula = '=arrayformula(substitute(split(substitute(A1:A'+numRows.toString()+',"|","^|"),"^"), "|",""))'  // We are splitting columns based on the | character
		spreadsheet.getRange('B1').setValue(formula);
	}
	all.getRange('L1').setValue("2. Getting List of Unique Interests")
	// Get unique list of interests
	var interest_sheet=SpreadsheetApp.openById(id).getSheetByName("interest");
	var end_data_row=interest_sheet.getLastRow()
	// Select all interests cells
	var interests=interest_sheet.getRange("B2:B"+end_data_row).getValues();
	// Convert to flattened array
	var interests_flat=[].concat.apply([], interests);
	var unique_interests=[]
	// For each cell if the interest isn't in the unique_interests array add it
	for (var row = 0; row < interests.length; row++) {
		if (interests_flat[row] != '' && unique_interests.indexOf(interests_flat[row]) == -1) {
			unique_interests.push(interests_flat[row]);
		}
	}
	all.getRange('L1').setValue("2b. Looking Up Interest Short Titles")
	var all_query='=UNIQUE(QUERY({';
	// Get Long Titles
	var long_titles=SpreadsheetApp.openById(id).getSheetByName("lookups").getRange("E2:E").getValues();
	var long_titles=[].concat.apply([], long_titles);
	// Get Short Titles
	var short_titles=SpreadsheetApp.openById(id).getSheetByName("lookups").getRange("F2:F").getValues();
	var short_titles=[].concat.apply([], short_titles);
	// For each intersts sheet
	for(var i=0; i < unique_interests.length; i++){  /// CHANGE FROM 2 BACK TO unique_interests.length!!!
		all.getRange('L1').setValue("3a. Adding Missing appl_ids for "+unique_interests[i])
		// Add each appl_id if it isn't already present
		// If we have a short interest name use it
		if(long_titles.indexOf(unique_interests[i])>-1){
			var interest_sheet=SpreadsheetApp.openById(id).getSheetByName(short_titles[long_titles.indexOf(unique_interests[i])]);
			all_query=all_query+"'"+short_titles[long_titles.indexOf(unique_interests[i])]+"'!A3:BA; "
		}
		// Else
		else{
			var interest_sheet=SpreadsheetApp.openById(id).getSheetByName(unique_interests[i]);
		}
		var interest_sheet_end_data_row=interest_sheet.getLastRow();
		// Start copying desired data
		var numApps=interest.getLastRow();
		// Add each appl_id if it isn't already present
		var existing_appl_ids = interest_sheet.getRange("G:G").getValues();
		var existing_flat=[].concat.apply([], existing_appl_ids);
		var appl_ids = interest.getRange("M2:M").getValues();
		var interest_cells=interest.getRange("B2:B").getValues();
		var interest_cells_flat=[].concat.apply([], interest_cells);
		var new_appl_ids=[]
		for (var row = 0; row < appl_ids.length; row++) {
			if (appl_ids[row] != '' && existing_flat.indexOf(Number(appl_ids[row]))==-1 && interest_cells_flat[row].trim()==unique_interests[i].trim()) {
				new_appl_ids.push(appl_ids[row]);
			}
		}
		var new_appl_ids_row=interest_sheet.getLastRow()+1;
		var new_appl_ids_cell="G"+(interest_sheet.getLastRow()+1)+":G"+(interest_sheet.getLastRow()+new_appl_ids.length)
		var new_appl_ids_range = interest_sheet.getRange(new_appl_ids_cell);
		if(new_appl_ids.length > 0)
			new_appl_ids_range.setValues(new_appl_ids);
		var end_data_row=interest_sheet.getLastRow()
		all.getRange('L1').setValue("3b. Beginning Simple Imports for for "+unique_interests[i])
		var import_fields= [ 
		["H","main","BN","B"],     // Appl Complete
		["I","main","D","B"],      // Last Name
		["J","main","C","B"],      // First Name
		["K","main","Q","B"],      // Email
		["O","option","B","A"],    // Non-Thesis Interest
		["P","main","AZ","B"],     // GPA Recent
		["AH","main","BE","B"],    // System Key
		["AI","interest","B","M"], // Interest
		["AJ","main","BT","B"],    // Ethnicity
		["AK","main","BU","B"],    // Gender
		["AL","main","BE","B"],	   // System Key		
		] // Ending Imports
		for (var row = 0; row < import_fields.length; row++){
			var main_cells=[]
			var formula=[]
			all.getRange('L1').setValue("3c. Running Simple Imports for for "+unique_interests[i]+' column '+import_fields[row][0])
			var import_sheet=import_fields[row][1]
			var import_column=import_fields[row][2]
			var import_id=import_fields[row][3]
			var all_column=import_fields[row][0]
			var main_range_cells=all_column+start_data_row+":"+all_column+end_data_row
			interest_sheet.getRange(main_range_cells).setValue("=IFERROR(INDEX("+import_sheet+"!"+import_column+":"+import_column+',MATCH(TRIM(INDIRECT(CONCAT("G",ROW()))),'+import_sheet+"!$"+import_id+":$"+import_id+",0)))");
		}
		// Degrees and tests require two lookups, one based on appl_id and one based on degree level
		// Column on all sheet, sheet to get data from, column to get data from, column of appl_id, degree level, degree level column
		var import_fields= [ 
		["Q","degree","L","X","Bachelor","O"],   // Bachelor Major GPA
		["R","degree","M","X","Bachelor","O"],   // Bachelor Overall GPA
		["S","degree","C","X","Bachelor","O"],   // Bachelor Institution
		["T","degree","O","X","Bachelor","O"],   // Bachelor Degree
		["U","degree","J","X","Bachelor","O"],   // Bachelor Date
		["V","degree","L","X","Master","O"],     // MS Major GPA
		["W","degree","M","X","Master","O"],     // MS Overall GPA
		["X","degree","C","X","Master","O"],     // MS Institution
		["Y","degree","O","X","Master","O"],     // MS Degree 
		["Z","degree","J","X","Master","O"],     // MS Degree Date
		["AA","test","G","B","TOEFLI","D"],		 // TOEFLI Score
		["AB","test","G","B","GRE  V","D"],		 // GRE V Score
		["AC","test","F","B","GRE  V","D"],		 // GRE V Percentile
		["AD","test","G","B","GRE Q ","D"],		 // GRE Q Score
		["AE","test","F","B","GRE Q ","D"],		 // GRE Q Percentile
		["AF","test","G","B","GREW  ","D"],		 // GRE W Score
		["AG","test","F","B","GREW  ","D"],		 // GRE Q Percentile
		]  // Ending Imports
		for (var row = 0; row < import_fields.length; row++){
			all.getRange('L1').setValue("3d. Running Complex Imports for for "+unique_interests[i]+' column '+import_fields[row][0])
			var main_cells=[]
			var formula=[]
			var import_sheet=import_fields[row][1]
			var import_column=import_fields[row][2]
			var import_id=import_fields[row][3]
			var import_filter=import_fields[row][4]
			var import_filter_column=import_fields[row][5]
			var all_column=import_fields[row][0]
			// Since Bachelor and Master degrees can have different names we just search for the first leter
			if(import_sheet=='degree' && import_filter=='Bachelor')
				import_filter="LEFT(degree!"+import_filter_column+":"+import_filter_column+')="B"';
			else if(import_sheet=='degree' && import_filter=='Master')
				import_filter="LEFT(degree!"+import_filter_column+":"+import_filter_column+')="M"';
			else
				import_filter=import_sheet+"!"+import_filter_column+":"+import_filter_column+'="'+import_filter+'"';
			// Annoyingly the test sheet doesn't use appl_id, rather it uses system_key so this changes the lookup column
			if(import_sheet=='test')
				all_id="AL"
			else
				all_id="G"
			var main_range_cells=all_column+start_data_row+":"+all_column+end_data_row
			interest_sheet.getRange(main_range_cells).setValue("=IFERROR(INDEX("+import_sheet+"!"+import_column+":"+import_column+",MATCH(1,("+import_sheet+"!"+import_id+":"+import_id+"=TRIM(INDIRECT(CONCAT(\""+all_id+"\",ROW()))))*("+import_filter+"),0),0))");
		}
		// Resident Status and Degree Goals require lookups for actual values
		var import_fields = [
		["L","main","BM","B","$A$2:$B$8"],   // Resident Status
		["M","main","J","B","$C$2:$D$8"],   // Degree Goal
		["N","main","I","B","$C$2:$D$8"],   // Ultimate Degree Goal
		] // Ending lookups
		for (var row = 0; row < import_fields.length; row++){
			all.getRange('L1').setValue("3e. Running Lookups for  "+unique_interests[i]+' column '+import_fields[row][0])
			var main_cells=[]
			var formula=[]
			var import_sheet=import_fields[row][1]
			var import_column=import_fields[row][2]
			var import_id=import_fields[row][3]
			var all_column=import_fields[row][0]
			var lookup_range=import_fields[row][4]
			var main_range_cells=all_column+start_data_row+":"+all_column+end_data_row
			interest_sheet.getRange(main_range_cells).setValue("=VLOOKUP(INDEX("+import_sheet+"!"+import_column+":"+import_column+',MATCH(TRIM(INDIRECT(CONCAT("G",ROW()))),'+import_sheet+"!$"+import_id+":$"+import_id+",0)),lookups!"+lookup_range+",2,FALSE)");
		}
		// Removing all of the #N/A errors for not found values, also converts formulas to static values
		all.getRange('L1').setValue("3f. Converting Formulas to Values for "+unique_interests[i]);
		var values = interest_sheet.getDataRange().getValues();
		/*
		for(var row in values){
			var replaced_values = values[row].map(function(original_value){
				return original_value.toString().replace("#N/A","");
			});
			values[row] = replaced_values;
		}
		*/	
		interest_sheet.getDataRange().setValues(values);
	}
	all.getRange('L1').setValue("4. Populating All Sheet")
	all_query=all_query.substring(0,all_query.length-2)
	all_query=all_query+'}, "SELECT * WHERE Col7 IS NOT NULL ORDER BY Col9, Col10"))'
	all.getRange('A3').setValue(all_query)
	all.getRange('L1').setValue("Processing Complete")
}