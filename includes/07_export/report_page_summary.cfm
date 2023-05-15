<cfif #get_Summary.recordcount# GT 0>

		<cfoutput query="get_Summary" group="PropNum">
			<table width="100%" cellpadding="2" cellspacing="0" border="2">
				<tr>
					<td class="navbottom" colspan="5">
						<span style="margin-left: 10px;"><strong>#PropNum# - #PropName#</strong></span>
					</td>
				</tr>
				<cfoutput group="MainCat">
					<tr>
						<td class="navbar" style="border-top: solid 1px;"  colspan="5" valign="top">
							<span style="margin-left: 25px;"><strong>#MainCat#</strong></span>
						</td>
					</tr>
					<cfoutput group="item_Name">
						<tr>
							<td></td>
							<td colspan="4" valign="top">	
								<strong>#item_Name#</strong><br/>
								<cfoutput group="keyName">
									<span ><i><strong>#keyName#</strong></i></span>
									<ul style="margin-top: 0px;">	
										<cfoutput group="keyValue">
											<cfif #WO_keyName# EQ 'Deficiency'>
												<li><span><font color="red">#keyValue# (#iCount#) - Deficiency</font></span></li>
											<cfelseif #WO_keyName# EQ 'Work Order'>
												<li><span><font color="red"><strong>#keyValue# (#iCount#) - Work Order</strong></font></span></li>
											<cfelse>
												<li><span>#keyValue# (#iCount#)</span></li>
											</cfif>
										</cfoutput>
									</ul>
								</cfoutput>		
							</td>
						</tr>
					</cfoutput>
				</cfoutput>		
			</table>
		</cfoutput>			
		
   
<script language="javascript">
	
function submitExport(sType) {
	document.getElementById("exportType").value = sType; 
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


				
				
				