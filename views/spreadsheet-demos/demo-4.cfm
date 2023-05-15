
<cfscript>
	sPath = ExpandPath("*.*");
	sDirectory = getDirectoryFromPath(sPath);

	csvFile = fileread("#sDirectory#/includes/csv/salesChart.csv");
	spreadsheet = wirebox.getInstance( "Spreadsheet@spreadsheet-cfml" );
	csvData = spreadsheet.csvToQuery( csv="#csvFile#", firstRowIsHeader=true );
</cfscript>

<cfif isdefined('FORM.submit')>

		<cfscript>
			sObj = spreadsheet.read("#ExpandPath( "./" )#/template_files/DynamicChartDemo.xlsx");

            spreadsheet.SetActiveSheet(sObj, "ChartData");
            
            spreadsheet.SetCellValue(sObj, "Bonuses", 1, 2)
            for(i=1;i<=csvData.recordCount;i++){
                spreadsheet.SetCellValue(sObj, "#csvData.Associate[i]#", i+1, 1);
                spreadsheet.SetCellValue(sObj, #csvData.sales[i]#, i+1, 2)
            }

			startCellColRef="A";
			startCellRowRef="2";
			endCellColRef="A";
			endCellRowRef="#(csvData.recordCount+1)#";
			sheetName = "ChartData";
            NameReference = sObj.getName("Associates");
            referenceString = sheetName&"!$"&startCellColRef&"$"&startCellRowRef&":$"&endCellColRef&"$"&endCellRowRef;
			NameReference.setRefersToFormula(referenceString);
			
			startCellColRef="B";
			startCellRowRef="2";
			endCellColRef="B";
			endCellRowRef="#(csvData.recordCount+1)#";
			sheetName = "ChartData";
            NameReference = sObj.getName("Sales");
            referenceString = sheetName&"!$"&startCellColRef&"$"&startCellRowRef&":$"&endCellColRef&"$"&endCellRowRef;
			NameReference.setRefersToFormula(referenceString);
			       
            spreadsheet.SetActiveSheet(sObj, 'DynamicChart');
        
        </cfscript>

    <cfheader name="Content-Disposition" value="inline; filename=chart.xlsx">
	<cfcontent type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" variable="#spreadSheet.ReadBinary(sObj)#" />

</cfif>

<div class="text-center card shadow-sm bg-light border border-5 border-white">
	<div class="card-body">

		<h1 class="display-5 fw-bold">
			<cfoutput>#prc.welcomeMessage#</cfoutput>
		</h1>

		<div class="col-lg-12 mx-auto">
			<p class="lead mb-4">
			
			<form name="frmSales" action="" method="post">
				<input name="submit" type="Submit" class="btn btn-dark btn-lg" role="button" value="Create Excel" class="btn" /><br/><br/>
			</form>

		</div>
	</div>
</div>

