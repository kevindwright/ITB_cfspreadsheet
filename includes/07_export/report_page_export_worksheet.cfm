
<cfscript>
get_PropInfo = queryNew("building,Descr,address,city,state,postalCode,regional",
                        "integer,varchar,varchar,varchar,varchar,varchar,varchar",
                        {
                            "building":8657,
                            "descr":"Coco Village",
                            "address":"1200 Bichon Frise",
                            "city":"Barker",
                            "state":"CA",
                            "postalCode":"92807",
                            "regional":"Kevin D. Wright"
                        });
                        				
</cfscript>

    
    <cfif #get_Summary.recordcount# GT 0>


        <cfscript>
            theFile = "#GetDirectoryFromPath(GetCurrentTemplatePath())#worksheet.xls"
            spreadsheet = wirebox.getInstance( "Spreadsheet@spreadsheet-cfml" );
        
            endrow = spreadsheet.read(src="#theFile#", format="query", sheetNumber="3", rows="1", columns="15");

            //writedump(var="#endrow#");
            //abort;

            ss_schema = spreadsheet.read(src="#theFile#", format="query", headerRow="1", sheetNumber="3", rows="2-#((endrow.column1)-1)#", columns="1-14");

            objSpreadSheet = spreadSheet.Read("#GetDirectoryFromPath(GetCurrentTemplatePath())#/worksheet.xls");
        </cfscript>

        <cfif get_PropInfo.recordcount GT 0 >
            <cfscript>
                spreadsheet.SetActiveSheet (objSpreadSheet, "UBU Summary pg 1");
                spreadsheet.SetCellValue(objSpreadSheet, #get_PropInfo.descr#, 2, 15);
                spreadsheet.SetCellValue(objSpreadSheet, #dateformat(now(),"mm/dd/yyyy")#, 3, 15);
                spreadsheet.SetCellValue(objSpreadSheet, #get_PropInfo.regional#, 4, 15);
            </cfscript>
        </cfif>

        <cfoutput query="get_Summary">

            <cfquery dbtype="query" name="cell_schema">
                SELECT SS_Sheet, SS_Cell, SS_Col, SS_Row
                FROM ss_schema WHERE SectionVal = '#Trim(get_Summary.mainCat)#'
                        AND ItemVal = '#Trim(get_Summary.item_name)#'
                        AND keyNameVal = '#Trim(get_Summary.keyName)#'
                        AND KeyVal = '#Trim(get_Summary.keyValue)#' 
            </cfquery>

            <cfif cell_schema.recordcount GT 0 >
                <cfscript>
                    if (cell_schema.SS_Sheet == 1){
                        sheetName = "UBU Summary pg 1";
                    }
                    else{
                        sheetName = "UBU Summary pg 2";
                    }
                    spreadsheet.SetActiveSheet (objSpreadSheet, "#SheetName#");
                    spreadsheet.SetCellValue(objSpreadSheet, #iCount#, #cell_schema.SS_Row#, #cell_schema.SS_Col#,"numeric");
                </cfscript>

            </cfif>

        </cfoutput>

        <cfset cookie.userlogin = "">

        <cfif cookie.userlogin EQ "kwright"> 
            <cfscript>
                spreadsheet.CreateSheet (objSpreadSheet, "raw_data");
                spreadsheet.SetActiveSheet (objSpreadSheet, "raw_data");
                spreadsheet.AddRows(objSpreadSheet, get_Summary);
            </cfscript>
        <cfelse>
            <cfscript>
                spreadsheet.RemoveSheet (objSpreadSheet, "SchemaSheet");
                //spreadsheet.RemoveSheet (objSpreadSheet, "raw_data");
            </cfscript> 
        </cfif>

        <cfscript>
            //wb1 = objSpreadSheet.getWorkBook();

            // get a handle on the sheet object then then workbook
            wb = objSpreadSheet.getSheetAt(objSpreadSheet.getActiveSheetIndex()).getWorkBook();
            wb.setForceFormulaRecalculation(true);

            //writedump(#wb1#);
            //writedump(#wb#);
            //abort;

        </cfscript>

        <cfheader name="Content-Disposition" value="attachment; filename=summary_excel.xls">
        <cfcontent type="application/vnd.ms-excel" variable="#spreadSheet.ReadBinary(objSpreadSheet)#">

    <cfelse>
        
        <table width="100%" cellpadding="2" cellspacing="0" border="0">
			<tr>
				<td colspan="6" class="textImportant" align="center">
					<p>Your search returned zero records.</p>
				</td>	
			</tr>
		</table>

    </cfif>




				
				
				