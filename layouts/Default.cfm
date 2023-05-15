<cfoutput>
<!doctype html>
<html lang="en">
<head>
	<!--- Metatags --->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="IntoTheBox 2023 - Spreadsheet Magic">
    <meta name="author" content="Kevin D. Wright - Kinetic InterActive">

	<!---Base URL --->
	<base href="#event.getHTMLBaseURL()#" />

	<!---css --->

    <style>
        .bg-dark{ background-color: darkorange !important; }
        .text-blue { color:##379BC1; }
    </style>

	<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">
	<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.8.1/font/bootstrap-icons.css">

	<!--- Title --->
	<title>IntoTheBox 2023 - Spreadsheet Magic</title>
</head>
<body
	data-spy="scroll"
	data-target=".navbar"
	data-offset="50"
	style="padding-top: 60px"
	class="d-flex flex-column h-100"
>
	<!---Top NavBar --->
	<header>
		<nav class="navbar navbar-expand-lg navbar-dark bg-dark fixed-top">
			<div class="container-fluid">
				<!---Brand --->
				<a class="navbar-brand" href="#event.buildLink( 'main' )#">
					<strong><i class="bi bi-filetype-xlsx" style="margin-right:10px;"></i> Spreadsheet Magic</strong>
				</a>

				<!--- Mobile Toggler --->
				<button
					class="navbar-toggler"
					type="button"
					data-bs-toggle="collapse"
					data-bs-target="##navbarSupportedContent"
					aria-controls="navbarSupportedContent"
					aria-expanded="false"
					aria-label="Toggle navigation"
				>
					<span class="navbar-toggler-icon"></span>
				</button>

				<div class="collapse navbar-collapse" id="navbarSupportedContent">

					<!---About --->
					<ul class="navbar-nav ms-5 mb-2 mb-lg-0">
						<li class="nav-item dropdown">
							<a class="nav-link dropdown-toggle"
								href="##"
								id="navbarDropdown"
								role="button"
								data-bs-toggle="dropdown"
								aria-expanded="false">
								About <b class="caret"></b>
							</a>
							<ul class="dropdown-menu" aria-labelledby="navbarDropdown">
								<li>
									<a href="https://www.linkedin.com/in/kevindwright02/" target="_blank" class="dropdown-item">
										<i class="bi bi-linkedin"></i> LinkedIn
									</a>
								</li>
								<li>
									<a href="mailto:kevin@kinetic-interactive.com" class="dropdown-item">
										<i class="bi bi-envelope-fill"></i> Email Me
									</a>
								</li>
								<li>
									<a href="https://www.kinetic-interactive.com" target="_blank" class="dropdown-item">
										<i class="bi bi-star"></i>  Website
									</a>
								</li>

							</ul>
						</li>
					</ul>

					<!--- Community --->
					<ul class="navbar-nav mb-2 mb-lg-0">
						<li class="nav-item dropdown">
							<a class="nav-link dropdown-toggle"
								href="##"
								id="navbarDropdown"
								role="button"
								data-bs-toggle="dropdown"
								aria-expanded="false">
								Demos <b class="caret"></b>
							</a>
							<ul class="dropdown-menu" aria-labelledby="navbarDropdown">
								<li>
									<a class="dropdown-item" href="/spreadsheet-demos/demo1">Demo 1 - Formulas</a>
								</li>
								<li>
									<a class="dropdown-item" href="/spreadsheet-demos/demo2">Demo 2 - Formatting</a>
								</li>
								<li>
									<a class="dropdown-item" href="/spreadsheet-demos/demo3">Demo 3 - Conditional Formatting</a>
								</li>
								<li>
									<a class="dropdown-item" href="/spreadsheet-demos/demo4">Demo 4 - Dynamic Chart</a>
								</li>
                                <li>
									<hr class="dropdown-divider">
								</li>
                                <li>
                                    <div style="margin-left:10px;"><strong> Real World Examples </strong></div>
                                </li>
                                <li>
									<hr class="dropdown-divider">
								</li>
                                <li>
									<a class="dropdown-item" href="/spreadsheet-demos/demo5">Demo 5 - Importing Data</a>
								</li>
                                <li>
									<a class="dropdown-item" href="/spreadsheet-demos/demo6">Demo 6 - Exporting Data</a>
								</li>
                                <li>
									<a class="dropdown-item" href="/spreadsheet-demos/demo7">Demo 7 - Bonus Demo</a>
								</li>
							</ul>
						</li>
					</ul>
				</div>
			</div>
		</nav>
	</header>

	<!---Container And Views --->
	<main class="flex-shrink-0">
		#renderView()#
	</main>

	<!--- Footer --->
	<footer class="w-100 bottom-0 position-fixed border-top py-3 mt-5 bg-light">
		<div class="container">
			<p class="float-end">
				<a href="https://www.kinetic-interactive.com" >
					<a href="https://www.kinetic-interactive.com"><img src="https://www.kinetic-interactive.com/assets/images/kia_logo.png" width="200"/>
				</a>
			</p>

		</div>
	</footer>

	<!---js --->
	<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>
	<script defer src="https://unpkg.com/alpinejs@3.x.x/dist/cdn.min.js"></script>


</body>
</html>
</cfoutput>
