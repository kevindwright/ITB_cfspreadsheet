
<cfscript>
	sPath = ExpandPath("*.*");
	sDirectory = getDirectoryFromPath(sPath);

	csvFile = fileread("#sDirectory#/includes/csv/sales.csv");
	spreadsheet = wirebox.getInstance( "Spreadsheet@spreadsheet-cfml" );
	csvData = spreadsheet.csvToQuery( csv="#csvFile#", firstRowIsHeader=true );
</cfscript>

<cfif isdefined('FORM.submit')>
    
		<cfscript>

			public void function formatCell(iValue,iRow,iColumn,nThreshold){
				if (#iValue# GTE #nThreshold#){
					spreadsheet.FormatCell(sObj,format1,iRow,iColumn);
				}
			}

			format1=StructNew(); 
			format1.font="serif"; 
			format1.fontsize="12"; 
			format1.color="white"; 
			format1.bold="true"; 
			format1.alignment="center";
			format1.fgcolor = "maroon"; 

			iLimit = 9000;
			sObj = spreadsheet.New("Sales") ;
			spreadsheet.addrow(sObj,"SalesRep,District,Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec");

			cfloop(query = csvData) {
				spreadsheet.addrow(sObj,"#SalesRep#,#District#,#Jan#,#Feb#,#Mar#,#Apr#,#May#,#Jun#,#Jul#,#Aug#,#Sep#,#Oct#,#Nov#,#Dec#");
				formatCell(#Jan#,(#currentRow#+1),3,#iLimit#);
				formatCell(#Feb#,(#currentRow#+1),4,#iLimit#);
				formatCell(#Mar#,(#currentRow#+1),5,#iLimit#);
				formatCell(#Apr#,(#currentRow#+1),6,#iLimit#);
				formatCell(#May#,(#currentRow#+1),7,#iLimit#);
				formatCell(#Jun#,(#currentRow#+1),8,#iLimit#);
				formatCell(#Jul#,(#currentRow#+1),9,#iLimit#);
				formatCell(#Aug#,(#currentRow#+1),10,#iLimit#);
				formatCell(#Sep#,(#currentRow#+1),11,#iLimit#);
				formatCell(#Oct#,(#currentRow#+1),12,#iLimit#);
				formatCell(#Nov#,(#currentRow#+1),13,#iLimit#);
				formatCell(#Dec#,(#currentRow#+1),14,#iLimit#);
			}

			format2 = StructNew();
			format2.dataformat = "$######,####0.00";
		
			spreadsheet.FormatColumn(sObj,format2,14) ;
			spreadsheet.FormatRow(sObj, {bold=TRUE, alignment="center", fgcolor="grey_50_percent", color="white"}, 1) ;
			
			format3 = StructNew() ;
			format3.bold = "true" ;
			format3.bottomborder = "thin" ;
			format3.bottombordercolor = "black" ;
			frmat3.fgcolor = "light_yellow" ;
			
			spreadsheet.FormatCellRange(sObj, format3, 4, 1, 4, 11) ;

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