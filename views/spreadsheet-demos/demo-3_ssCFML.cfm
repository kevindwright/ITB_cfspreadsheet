<cfscript>
	sPath = ExpandPath("*.*");
	sDirectory = getDirectoryFromPath(sPath);

	csvFile = fileread("#sDirectory#/includes/csv/sales.csv");
	spreadsheet = wirebox.getInstance( "Spreadsheet@spreadsheet-cfml" );
	csvData = spreadsheet.csvToQuery( csv="#csvFile#", firstRowIsHeader=true );
</cfscript>

<cfif isdefined('FORM.submit')>

<cfset sObj = spreadsheet.New("Sales") /> 
        
	<cfscript>
	
		spreadsheet.addrow(sObj,"SalesRep,District,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec,Total,Avg"); 

		cfloop(query = csvData) {
			spreadsheet.addrow(sObj,"#SalesRep#,#District#,#Jan#,#Feb#,#Mar#,#Apr#,#May#,#Jun#,#Jul#,#Aug#,#Sep#,#Oct#,#Nov#,#Dec#");
			spreadsheet.SetCellFormula(sObj, "SUM(C#currentRow+1#:N#currentRow+1#)", #currentRow#+1, 15);
			spreadsheet.SetCellFormula(sObj, "O#currentRow+1#/12", #currentRow#+1, 16);
		}

		rowDataStart=2;
		rowDataEnd=csvData.recordCount+1;
		
		spreadsheet.SetCellValue(sObj, "Total:", rowDataEnd+1, 2);
		spreadsheet.SetCellFormula(sObj, "SUM(C#rowDataStart#:C#rowDataEnd#)", rowDataEnd+1, 3);
		spreadsheet.SetCellFormula(sObj, "SUM(D#rowDataStart#:D#rowDataEnd#)", rowDataEnd+1, 4);
		spreadsheet.SetCellFormula(sObj, "SUM(E#rowDataStart#:E#rowDataEnd#)", rowDataEnd+1, 5);
		spreadsheet.SetCellFormula(sObj, "SUM(F#rowDataStart#:F#rowDataEnd#)", rowDataEnd+1, 6);
		spreadsheet.SetCellFormula(sObj, "SUM(G#rowDataStart#:G#rowDataEnd#)", rowDataEnd+1, 7);
		spreadsheet.SetCellFormula(sObj, "SUM(H#rowDataStart#:H#rowDataEnd#)", rowDataEnd+1, 8);
		spreadsheet.SetCellFormula(sObj, "SUM(I#rowDataStart#:I#rowDataEnd#)", rowDataEnd+1, 9);
		spreadsheet.SetCellFormula(sObj, "SUM(J#rowDataStart#:J#rowDataEnd#)", rowDataEnd+1, 10);
		spreadsheet.SetCellFormula(sObj, "SUM(K#rowDataStart#:K#rowDataEnd#)", rowDataEnd+1, 11);
		spreadsheet.SetCellFormula(sObj, "SUM(L#rowDataStart#:L#rowDataEnd#)", rowDataEnd+1, 12);
		spreadsheet.SetCellFormula(sObj, "SUM(M#rowDataStart#:M#rowDataEnd#)", rowDataEnd+1, 13);
		spreadsheet.SetCellFormula(sObj, "SUM(N#rowDataStart#:N#rowDataEnd#)", rowDataEnd+1, 14);
		spreadsheet.SetCellFormula(sObj, "SUM(M#rowDataStart#:M#rowDataEnd#)", rowDataEnd+1, 15);
		spreadsheet.FormatCell(sObj, {bold=TRUE, alignment="center", fgcolor="red", color="white"}, rowDataEnd+1,15);

		for (i=3; i <= 16; i++) {
			spreadsheet.FormatColumn(sObj,{dataformat = "$######,####0.00"},i);
		}

		spreadsheet.FormatRow(sObj, {bold=TRUE, alignment="center", fgcolor="grey_50_percent", color="white"}, 1); 
		spreadsheet.FormatRow(sObj, {bold=TRUE}, rowDataEnd+1); 

		// get a handle on the sheet object
		objSheet = sObj.getSheetAt(sObj.getActiveSheetIndex());

		// get a handle on the conditional formatting property of the sheet
		objCondFormat = objSheet.getSheetConditionalFormatting();

		// create a formatting rule
		ConditionalFormattingRule = objCondFormat.createConditionalFormattingRule(1,"6000","7000");
		//ConditionalFormattingRule = objCondFormat.createConditionalFormattingRule(5,"5000");
		
		// create a format for the rule
		PatternFormatting = ConditionalFormattingRule.createPatternFormatting();
		PatternFormatting.setFillBackgroundColor(25); //54
		FontFormatting = ConditionalFormattingRule.createFontFormatting();
		FontFormatting.setFontColorIndex(1);
		//PatternFormatting.setFillBackgroundColor(4);

		regions = arrayNew(1);
		//regions[1] = createObject('java', 'org.apache.poi.ss.util.CellRangeAddress').init(1,100,1,15);
		regions[1] = spreadsheet.getRangeHelper().getCellRangeAddressFromColumnAndRowIndices( { startRow:1, endRow:rowDataEnd, startColumn: 2, endColumn: 15} );
		
		// writedump(#objCondFormat#);
		 //writedump(#ConditionalFormattingRule#);
		 //writedump(#FontFormatting#);

		 //writedump(spreadsheet);
		 ///writedump(spreadsheet.getRangeHelper());
		 //abort;
 
		//Apply Conditional Formatting rule defined above to the regions 
		objCondFormat.addConditionalFormatting(regions, ConditionalFormattingRule);

		spreadsheet.write( sObj, "#ExpandPath( 'tmpfiles\' )#sales.xls", true );

	</cfscript>


    <cfheader name="Content-Disposition" value="inline; filename=sales.xls">
	<cfcontent type="application/vnd.ms-excel" file="#ExpandPath( 'tmpFiles\' )#sales.xls">



</cfif>

<div class="text-center card shadow-sm bg-light border border-5 border-white">
	<div class="card-body">

		<h1 class="display-5 fw-bold">
			<cfoutput>#prc.welcomeMessage#</cfoutput>
		</h1>

		<div class="col-lg-12 mx-auto">
			<p class="lead mb-4">
				<table id="demo" class="display" style="width:100%;" border=0>
					<thead>
						<tr>
							<th>Sales Rep</th>
							<th>District</th>
							<th>Jan</th>
							<th>Feb</th>
							<th>Mar</th>
							<th>Apr</th>
							<th>May</th>
							<th>Jun</th>
							<th>Jul</th>
							<th>Aug</th>
							<th>Sep</th>
							<th>Oct</th>
							<th>Nov</th>
							<th>Dec</th>
						</tr>
					</thead>
					<tbody>
						<cfoutput query="csvData"> 
						<tr>
							<td>#SalesRep#</td>
							<td>#District#</td>
							<td>#Jan#</td>
							<td>#Feb#</td>				
							<td>#Mar#</td>
							<td>#Apr#</td>
							<td>#May#</td>
							<td>#Jun#</td>
							<td>#Jul#</td>				
							<td>#Aug#</td>
							<td>#Sep#</td>
							<td>#Oct#</td>
							<td>#Nov#</td>
							<td>#Dec#</td> 
						</tr>   
						</cfoutput>
					</tbody>
				</table>
			</p>

			<form name="frmSales" action="" method="post">
				<input name="submit" type="Submit" class="btn btn-dark btn-lg" role="button" value="Export To Excel" class="btn" /><br/><br/>
			</form>

		</div>
	</div>
</div>

<link href="https://cdn.datatables.net/1.10.16/css/jquery.dataTables.min.css" rel="stylesheet" type="text/css" media="screen" />

