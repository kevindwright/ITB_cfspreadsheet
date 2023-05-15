

<!--- Upload the file that has just been submitted --->
<cfif IsDefined("FORM.fileName") and FORM.fileName NEQ "">







<cfelse>



<div class="text-center card shadow-sm bg-light border border-5 border-white">
	<div class="card-body">

		<h1 class="display-5 fw-bold">
			<cfoutput>#prc.welcomeMessage#</cfoutput>
		</h1>

		<div class="col-lg-12 mx-auto">
			<p class="lead mb-4">
			
			<form action="" method="post" enctype="multipart/form-data" name="frmUpload" id="frmUpload">
                <input type="File" name="fileName" class="fmenu">
                &nbsp;&nbsp;&nbsp;&nbsp;
                <input type="Submit" name="btnUpload" value="Upload File" class="button">
            </form>

		</div>
	</div>
</div>

</cfif>