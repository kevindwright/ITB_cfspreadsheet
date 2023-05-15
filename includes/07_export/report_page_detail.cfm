
<cfif #get_Details.recordcount# GT 0>

	<table width="100%">
		<tr>
			<td>

				<cfoutput query="get_Details" group="PropNum">

					<table width="100%" cellpadding="2" cellspacing="0" border="0">
						<tr>
							<td>
								<table width="100%" cellpadding="2" cellspacing="0" border="0">
									<tr>
										<td class="navbottom" colspan="5">
											<strong>#PropNum# - #PropName#</strong>
										</td>
									</tr>
									<tr>
										<td class="navbar1" valign="top" colspan="4">
											<strong>Category / Item</strong>
										</td>
										<td  class="navbar1">
											<strong>Comments</strong>
										</td>
									</tr>

									<cfset iCounter = 0>
									<cfoutput group="unit_name">
											<cfset iCounter = iCounter + 1>
											<tr>
												<td colspan="5" class="RuleGrey"></td>
											</tr>
											<tr bgcolor="###iif(iCounter MOD 2,DE('ffffff'),DE('efefef'))#">
												<td colspan="5">
													<strong>Unit ## #unit_name#</strong>
												</td>
											</tr>
											<cfoutput group="sectionName">
												<tr bgcolor="###iif(iCounter MOD 2,DE('ffffff'),DE('efefef'))#">
													<td colspan="5" class="RuleGrey"></td>
												</tr>
												<tr bgcolor="###iif(iCounter MOD 2,DE('ffffff'),DE('efefef'))#">
													<td colspan="5">
														<strong>#sectionName#</strong>
													</td>
												</tr>

												<cfoutput group="item_name">
													<tr bgcolor="###iif(iCounter MOD 2,DE('ffffff'),DE('efefef'))#" >
														<td valign="top" colspan="2" width="25%">
															
														</td>
														<td valign="top">
															#item_name#
														</td>
														<td valign="top">
															#keyName# / #WO_key#=#WO_Value#
														</td>
														<td valign="top">
															#notes#
														</td>
													</tr>
												</cfoutput>
											</cfoutput>
											<tr>
												<td colspan="5" class="RuleGrey"></td>
											</tr>
									</cfoutput>
								</table>
							</td>
						</tr>
					</table>
				</cfoutput>
		
			</td>	
		</tr>
	</table>			
		<!--- <cfoutput>				
			<form name="frmExport" id="frmExport" action="/index.cfm?sPageInclude=happyCo/report_page.cfm" method="post">
				<input type="hidden" name="rdoRptType" value="#Variables.sRptType#" />
				<input type="hidden" name="Inspection" value="#Variables.sInspection#" />
				<input type="hidden" name="PropList" value="#Variables.sProperty#" />
				<input type="hidden" name="export" value="1" >
				<input type="hidden" name="bSubmit" value="export" >
			</form>
		</cfoutput>				 --->
			

		<script language="javascript">

		function submitExport() {
			document.getElementById("frmExport").submit(); 
		}

		</script>

	<cfelse>


		<table width="100%" cellpadding="2" cellspacing="0" border="0">
			<tr>
				<td colspan="6" class="textImportant" align="center">
					<p>Your search returned zero records.</p>
				</td>	
			</tr>
		</table>

	</cfif>

<style>
#body, html, div, table, fieldset {
    font-size: 100%;
    line-height: 1.4em;
}

.navbar1 {
    color: steelblue;
    font-family: verdana, sans-serif;
    text-decoration: none;
}

.navbottom {
    color: #ffffff;
    font-family: verdana, sans-serif;
    text-decoration: none;
    font-weight: bold;
    padding: 3px;
    border-bottom: 1px solid #b2d9ec;
    background-color: steelblue;
    empty-cells: show;
}

.RuleGrey {
    padding: 0;
    height: 1px;
    overflow: hidden;
    background-color: #CCCCCC;
    empty-cells: show;
}
</style>