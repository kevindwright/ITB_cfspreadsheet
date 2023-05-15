
<cfif isdefined("Form.rdoRptType")>
	<cfset Variables.sRptType = #Form.rdoRptType#>
<cfelse>
	<cfset Variables.sRptType = "1">
</cfif>


<cfoutput>
<div class="text-center card shadow-sm bg-light border border-5 border-white">
	<div class="card-body">
		

		<h1 class="display-5 fw-bold">
			#prc.welcomeMessage#
		</h1>

		<div class="col-lg-6 mx-auto">
			<form action="" name="frmInspection" method="post">
				<div class="row">
					<div class="col-lg-5 text-start"></div>
					<div class="col-lg-7 text-start"><input type="radio" value="1" name="rdoRptType" <cfif #Variables.sRptType# EQ 1>checked</cfif> /> Detail</div>
				</div>
				<div class="row">
					<div class="col-lg-5 text-start"></div>
					<div class="col-lg-7 text-start"><input type="radio" value="2" name="rdoRptType" <cfif #Variables.sRptType# EQ 2>checked</cfif>/> Summary</div>
				</div>
				<div class="row">
					<div class="col-lg-5 text-start"></div>
					<div class="col-lg-7 text-start"><input type="radio" value="3" name="rdoRptType" <cfif #Variables.sRptType# EQ 3>checked</cfif>/> Worksheet</div>
				</div>
				<div class="row">
					<div class="col-lg-5 text-start"></div>
					<div class="col-lg-7 text-start"><p style="margin-top:20px;"><input type="submit" name="submit" value="View Report" class="button"></p></div>
				</div>
			</form>
		</div>
		<div class="col-lg-6 mx-auto">
		
			<div class="col-lg-12 text-start">	
				<cfscript>
					if (isdefined("FORM.Submit") OR isdefined("FORM.bSubmit")){

						spreadsheet = wirebox.getInstance( "Spreadsheet@spreadsheet-cfml" );

						sPath = ExpandPath("*.*");
						sDirectory = getDirectoryFromPath(sPath);

						switch(#FORM.rdoRptType#){
							case "1":
								csvFile = fileread("#sDirectory#/includes/csv/inspectionDetail.csv");
								get_Details = spreadsheet.csvToQuery( csv="#csvFile#", firstRowIsHeader=true );
								include "/includes/07_export/report_page_detail.cfm";
								break;
							case "2":
								csvFile = fileread("#sDirectory#/includes/csv/inspectionSummary.csv");
								get_Summary = spreadsheet.csvToQuery( csv="#csvFile#", firstRowIsHeader=true );
								include "/includes/07_export/report_page_summary.cfm";
								break;
							case "3":
								csvFile = fileread("#sDirectory#/includes/csv/inspectionData.csv");
								get_Summary = spreadsheet.csvToQuery( csv="#csvFile#", firstRowIsHeader=true );
								include "/includes/07_export/report_page_export_worksheet.cfm";
								break;
							default:
								break;
						}
					}
				</cfscript> 
			</div>
		</div>
	</div>
</div>

</cfoutput>