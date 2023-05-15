<!---
	Copyright (c) 2012 by Western National Group, LLC.
	FILE NAME	voucher_import.cfm
	AUTHOR		kwright@wng.com
	PROJECT  	InSite - Electronic Invoicing
	DESCRIPTION	Allows user to upload a file(s) to create dept vouchers
	DEVELOPMENT HISTORY
	DATE		WHO		VERSION		DESCRIPTION
	=======		=====	========	===============================
	04-15-12	KDW		1.0			Original Development
--->

<LINK REL="StyleSheet" HREF="styles.css" type="text/css">

<!--- Upload the file that has just been submitted --->
<cfif IsDefined("FORM.fileName") and FORM.fileName NEQ "">

<!--- set validation values --->
<cfset bValid = 1>  
<cfset bIsValid = 1>
         
<cfset sMessage = ''>
<cfset sLatestVersion = 'v8.4'>

<!--- set CORP or AP --->
<cfset variables.dept = #FORM.dept#>

<!--- Default variable values --->
<cfparam name="fileName" default="" type="string">
<cfset iInvoiceTotal = 0>
<cfset iCountTotal = 0>
<cfset iVoucherCount = 0>

<!--- ***************************************************************     
                            FUNCTIONS                                     
      *************************************************************** --->

<!--- VALIDATION --->
<cffunction name="validateCHRG_results">
<cfargument name="QryName" type="string" required="yes">

	<cfoutput query="#QryName#">
    
    	<!--- validate number of charges --->
		<cfif #COLUMN4# NEQ '0' AND #COLUMN4# NEQ ''>
            <cfif #IsValid ("integer", COLUMN4)# EQ 0>
            	<cfset bValid = 0>
            </cfif>
        </cfif>
        
        <!--- validate charge amount--->
		<cfif #COLUMN5# NEQ '0' AND #COLUMN5# NEQ ''>
            <cfif #IsValid ("float", COLUMN5)# EQ 0>
            	<cfset bValid = 0>
            </cfif>
        </cfif>
        
        <!--- validate total charge --->
		<cfif #COLUMN6# NEQ '0' AND #COLUMN6# NEQ ''>
            <cfif #IsValid ("float", COLUMN6)# EQ 0>
            	<cfset bValid = 0>
            </cfif>
        </cfif>

    	<!--- validate account code --->
		<cfif #COLUMN6# NEQ '0' AND #COLUMN6# NEQ ''>
            <cfif #COLUMN7# EQ 0 OR #COLUMN7# EQ "">
            	<cfset bValid = 0>
            </cfif>
        </cfif>

    </cfoutput>
</cffunction>

<!--- OUTPUT --->      
<cffunction name="outputCHRG_results">
<cfargument name="QryName" type="string" required="yes">
    <cfsavecontent variable="sHTML">
    <tr>
    	<td colspan="6" style="font-weight:bold"><cfoutput>#QryName#</cfoutput></td>
    </tr>
	<cfset iGroupTotal = 0> 
    <cfset iGroupCount = 0> 
    <cfset iGroupRow = 0>
	<cfoutput query="#QryName#">
        <cfset bIsValid = 1>
		<cfif #COLUMN6# NEQ '0' AND #COLUMN6# NEQ ''>
			<!--- validate number of charges --->
            <cfif #COLUMN4# NEQ '0' AND #COLUMN4# NEQ ''>
                <cfif #IsValid ("integer", COLUMN4)# EQ 0>
                    <cfset bIsValid = 0>
                </cfif>
			</cfif>
            <!--- validate charge amount--->
            <cfif #COLUMN5# NEQ '0' AND #COLUMN5# NEQ ''>
                <cfif #IsValid ("float", COLUMN5)# EQ 0>
                    <cfset bIsValid = 0>
                </cfif>
            </cfif>
            <!--- validate total charge --->
            <cfif #COLUMN6# NEQ '0' AND #COLUMN6# NEQ ''>
                <cfif #IsValid ("float", COLUMN6)# EQ 0>
                    <cfset bIsValid = 0>
                </cfif>
            </cfif>
            <!--- validate account code --->
            <cfif #COLUMN6# NEQ '0' AND #COLUMN6# NEQ ''>
                <cfif #COLUMN7# EQ 0 OR #COLUMN7# EQ "">
                    <cfset bIsValid = 0>
                </cfif>
            </cfif>
        	<cfif bIsValid EQ 0>
            	<tr bgcolor="##FF9999">  
            <cfelse>
            	<tr bgcolor="#IIF(iGroupRow MOD 2, DE('FFFFFF'), DE('F2F3F7'))#">
            </cfif>
            
                <td align="left" valign="top">#COLUMN1#</td>
                <td align="left" valign="top">#COLUMN2#</td>
                <td align="left" valign="top">#sDescription.COLUMN1#-<br/>#COLUMN3#</td>
                <td align="right" valign="top">#COLUMN4#</td>
                <td align="right" valign="top"><cfif #IsValid ("float", COLUMN5)# EQ 1>#dollarformat(replace(COLUMN5,",",""))#<cfelse>#COLUMN5#</cfif></td>
                <td align="right" valign="top">#COLUMN6#</td>
                <td align="right" valign="top">#COLUMN7#</td>
            </tr>
            <cfif #IsValid ("float", COLUMN6)# EQ 1>
                <cfset iGroupTotal = iGroupTotal + #replace(COLUMN6,",","")#>
            <cfelse>
                <cfset iGroupTotal = iGroupTotal>
            </cfif>
            
            <cfif #IsValid ("float", COLUMN4)# EQ 1>
                <cfset iGroupCount = iGroupCount + #COLUMN4#>
            <cfelse>
                <cfset iGroupCount = iGroupCount>
            </cfif>

            <cfset iVoucherCount = iVoucherCount + 1>
            <cfset iGroupRow = iGroupRow + 1>
        </cfif>
    </cfoutput>

    <tr bgcolor="##DEDEDE">
        <td colspan="2" align="right"></td>
		<td align="right" style="font-weight:bold">Total <cfoutput>#QryName#</cfoutput></td>
		<td align="right" style="font-weight:bold"><cfoutput>#iGroupCount#</cfoutput></td>
		<td align="right"></td><td align="right" style="font-weight:bold"><cfoutput>#dollarformat(iGroupTotal)#</cfoutput></td>
		<td>&nbsp;</td>
    </tr>
    </cfsavecontent>
    
    <cfif iGroupTotal GT 0>
    	<cfset iInvoiceTotal = iInvoiceTotal + iGroupTotal>
        <cfset iCountTotal = iCountTotal + iGroupCount>
        
    	<cfoutput>#sHTML#</cfoutput>
    </cfif>
    
</cffunction>


<!--- create unique filename --->
<cfset variables.sTimeStamp = toString(DateFormat(now(),"yyyymmdd")) & '_' & toString(TimeFormat(now(),"HHmmss"))>
<cfset variables.sFileNameDynamic = variables.sTimeStamp & ".xlsx">

<!--- set the directory path --->
<cfset sPath = ExpandPath( "./") >
<cfset sFilePath = sPath & "tmpfiles\">

<cffile action="UPLOAD" filefield="FORM.fileName" destination="#Variables.sFilePath##variables.sFileNameDynamic#" result="upload" accept="application/octet-stream, application/vnd.ms-excel, application/vnd.ms-excel.sheet.macroenabled.12, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" nameconflict="OVERWRITE" attributes="normal">
<cfset theFile = upload.serverDirectory & "/" & upload.serverFile>

<cfscript>

    spreadsheet = wirebox.getInstance( "Spreadsheet@spreadsheet-cfml" );

    // set the query beginning / end rows 
    E10_rows = spreadsheet.read(src=#theFile#, format="query", sheetNumber="2", rows="5", columns="4,5");
    EOM_rows = spreadsheet.read(src=#theFile#, format="query", sheetNumber="2", rows="6", columns="4,5");
    WNPLLC_rows = spreadsheet.read(src=#theFile#, format="query", sheetNumber="2", rows="7", columns="4,5");
    WNPM_rows = spreadsheet.read(src=#theFile#, format="query", sheetNumber="2", rows="8", columns="4,5");
    MAG_rows = spreadsheet.read(src=#theFile#, format="query", sheetNumber="2", rows="9", columns="4,5");
    WNC_rows = spreadsheet.read(src=#theFile#, format="query", sheetNumber="2", rows="10", columns="4,5");
    RGS_rows = spreadsheet.read(src=#theFile#, format="query", sheetNumber="2", rows="11", columns="4,5");

    theCols = "1,2,3,7,9,11,12";

    // create query objects 
   
    //E10 
    E10 = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="#E10_rows.column1#-#E10_rows.column2#", columns="#theCols#");
    //EOM   
    EOM = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="#EOM_rows.column1#-#EOM_rows.column2#", columns="#theCols#");
    //WNP_CORP 
    WNPLLC = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="#WNPLLC_rows.column1#-#WNPLLC_rows.column2#", columns="#theCols#");  
    // WNPM 
    WNPM = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="#WNPM_rows.column1#-#WNPM_rows.column2#", columns="#theCols#");       
    //MAG   
    MAG = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="#MAG_rows.column1#-#MAG_rows.column2#", columns="#theCols#");     
    //WNC  
    WNC = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="#WNC_rows.column1#-#WNC_rows.column2#", columns="#theCols#"); 
    //RGS  
    RGS = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="#RGS_rows.column1#-#RGS_rows.column2#", columns="#theCols#");
 

    versionCheck = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="3", columns="14");
    dtInvoice = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="3", columns="2");
    sInvoiceNum = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="4", columns="2");
    iVendorNum = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="5", columns="2");
    sVendorName = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="6", columns="2");
    sDescription = spreadsheet.read(src=#theFile#, format="query", sheetNumber="1", rows="7", columns="2");


    // general validation
    if (versionCheck.column1 NEQ #sLatestVersion#){
    	sMessage = sMessage & '<li>Invoice Template (#versionCheck.column1#) is an incorrect version.<br/>Please get the updated version (#sLatestVersion#) before proceeding.</li>';
    }
	if (dtInvoice.column1 EQ ''){
    	sMessage = sMessage & '<li>Invoice date must be included.</li>';
    }
    if (sInvoiceNum.column1 EQ ''){
    	sMessage = sMessage & '<li>Invoice number must be included.</li>';
    }
    if (len(sInvoiceNum.column1) GTE 20){
    	sMessage = sMessage & '<li>Invoice number must be 20 characters or less.</li>';
    }
    if (iVendorNum.column1 EQ ''){
    	sMessage = sMessage & '<li>Vendor Number must be included.</li>';
    }
    if (sVendorName.column1 EQ ''){
    	sMessage = sMessage & '<li>Vendor Name must be included.</li>';
    }
    if (sDescription.column1 EQ ''){
    	sMessage = sMessage & '<li>Description must be included.</li>';
    }

    // line item validation
    if (variables.dept EQ 'ap'){
        // IMPORT PROPERTY CHARGES 
        validateCharges = validateCHRG_results('E10');
        validateCharges = validateCHRG_results('EOM');
    }else{
        // IMPORT INTERDEPARTMENT CHARGES
        validateCharges = validateCHRG_results('WNPLLC');
        validateCharges = validateCHRG_results('WNPM');
        validateCharges = validateCHRG_results('MAG');
        validateCharges = validateCHRG_results('WNC');
        validateCharges = validateCHRG_results('RGS');
    }
	
    if (bValid EQ 0){
    	sMessage = sMessage & '<li>HighlightedRows may have one or more issues<ul>';
        sMessage = sMessage & '<li>Number of Charges must be numeric value</li>';
        sMessage = sMessage & '<li>Charge Amount is invalid</li>';
        sMessage = sMessage & '<li>Account Code missing</li>';
        sMessage = sMessage & '</ul></li>';
    }

</cfscript>


<table width="70%" align="center">
    <cfif sMessage NEQ ''>
        <tr>
            <td>

                <style>
                table#errorbox{
                    background:#FFC0C0;
                    border:1px solid #C00000;
                    font-size:16px;
                    margin-bottom:18px;
                    padding:25px;
                    position:relative;
                    text-align:left;
                    width:100%; 
                    height:100%; 
                }
                img#alert{
                    width:100px;
                }
                </style>
                
                <cfoutput>
                    <p>&nbsp;</p> 
                    <table id="errorbox">
                        <tr>
                            <td width="25%" align="center">
                                <img id="alert" src="alert.png" />
                            </td>
                            <td width="75%" align="left">
                                <strong>There is a problem with the uploaded spreadsheet.<br/>
                                Please correct the following issues;</strong><br/>
                                <ul>#sMessage#</ul>
                            </td>
                        </tr>
                    </table> 
                    <p>&nbsp;</p>
                </cfoutput>   
            </td>
        </tr>
    </cfif>
        <tr>
            <td>    
           
            <cfoutput> 
                <p>&nbsp;</p> 
                <table width="35%" bgcolor="##DEDEDE">
                    <tr>
                        <td align="right"><strong>Invoice ##:</strong></td>
                        <td align="left" style="padding-left:20px;"> #sInvoiceNum.column1#</td>
                    </tr>
                    <tr>
                        <td align="right"><strong>Invoice Date:</strong></td>
                        <td align="left" style="padding-left:20px;"> #dateformat(dtInvoice.column1,'mm/dd/yyyy')#</td>
                    </tr>
                    <tr>
                        <td align="right"><strong>Vendor ##:</strong></td>
                        <td align="left" style="padding-left:20px;"> #iVendorNum.column1#</td>
                    </tr>
                    <tr>
                        <td align="right"><strong>Vendor Name:</strong></td>
                        <td align="left" style="padding-left:20px;"> #sVendorName.column1#</td>
                    </tr>
                </table>
                <p>&nbsp;</p>
            </cfoutput>

            <table width="100%">
                <tr>
                    <th>Building</th><th>Property</th><th>Description</th><th style="text-align:right"># of Charges</th><th style="text-align:right">Charge Amt.</th><th style="text-align:right">Tot. Charge</th><th style="text-align:right">Account Code</th>
                </tr>
                <cfif variables.dept EQ 'ap'>
                    <!--- IMPORT PROPERTY CHARGES --->
                    <cfset outPutHTML = outputCHRG_results('E10')>
                    <cfset outPutHTML = outputCHRG_results('EOM')>
                <cfelse>
                    <!--- IMPORT INTERDEPARTMENT CHARGES --->
                    <cfset outPutHTML = outputCHRG_results('WNPLLC')>
                    <cfset outPutHTML = outputCHRG_results('WNPM')>
                    <cfset outPutHTML = outputCHRG_results('MAG')>
                    <cfset outPutHTML = outputCHRG_results('WNC')>
                    <cfset outPutHTML = outputCHRG_results('RGS')>
                </cfif>

                <tr>
                    <td colspan="6">&nbsp;</td>
                </tr>
                <tr bgcolor="#DEDEDE">
                    <th></th><th style="font-weight:bold; text-align:right">Invoice Total</th><th style="font-weight:bold; text-align:right"><cfoutput>#iVoucherCount# Vouchers</cfoutput></th><th  style="font-weight:bold; text-align:right"><cfoutput>#iCountTotal#</cfoutput></th><th>&nbsp;</th><th style="font-weight:bold; text-align:right"><cfoutput>#dollarformat(iInvoiceTotal)#</cfoutput></th><th>&nbsp;</th>
                </tr>
                <tr>
                    <td colspan="6">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan="6" align="right">
                    
                    <cfif sMessage EQ ''>
                        <form name="" id="" action="create_vouchers.cfm" method="post">
                            <input type="hidden" name="docName" id="docName" value="<cfoutput>#theFile#</cfoutput>">
                            <input type="hidden" name="dept" id="dept" value="<cfoutput>#variables.dept#</cfoutput>">
                            <input type="submit" name="submit" id="submit" value="Create Vouchers" class="button">
                        </form>
                    </cfif>
                    </td>
                </tr>
            </table>

        </td>
    </tr>
</table>


<cfelse>


<!--- <cfset variables.dept = #URL.dept#> --->
<cfset variables.dept = "AP">

<script language="JavaScript1.2">
<!--
function CheckForm(){
	
	// Allowed file types
	var FileExt = new Array();
	FileExt[0]=".xls";
	var FileName = document.getElementById('fileName').value.toLowerCase();
	var found = 0;
	for(i=0;i<FileExt.length;i++){
		if(FileName.indexOf(FileExt[i]) > 0)
		found = 1;
	}
	if(found == 0){
		alert('The file type you have selected is not allowed.\nPlease choose only an XLS file.\n');
		return false;
	}
}
function CheckFormCSV(){
	
	// Allowed file types
	var FileExt = new Array();
	FileExt[0]=".csv";
	var FileName = document.getElementById('fileNameCSV').value.toLowerCase();
	var found = 0;
	for(i=0;i<FileExt.length;i++){
		if(FileName.indexOf(FileExt[i]) > 0)
		found = 1;
	}
	if(found == 0){
		alert('The file type you have selected is not allowed.\nPlease choose only an CSV file.\n');
		return false;
	}
}
//-->
</script>



<div class="text-center card shadow-sm bg-light border border-5 border-white">
	<div class="card-body">

		<h1 class="display-5 fw-bold">
			<cfoutput>#prc.welcomeMessage#</cfoutput>
		</h1>

		<div class="col-lg-12 mx-auto">
			<p class="lead mb-4">
			
			<form action="" method="post" enctype="multipart/form-data" name="frmUpload" id="frmUpload" onSubmit="return CheckForm()">
                <input type="File" name="fileName" class="fmenu">
                <input type="hidden" name="dept" id="dept" value="<cfoutput>#variables.dept#</cfoutput>">
                &nbsp;&nbsp;&nbsp;&nbsp;
                <input type="Submit" name="btnUpload" value="Upload File" class="button">
            </form>

		</div>
	</div>
</div>

</cfif>