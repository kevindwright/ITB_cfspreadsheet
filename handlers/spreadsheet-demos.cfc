component extends="coldbox.system.EventHandler" {

	/**
	 * Default Action
	 */
	function index( event, rc, prc ) {
		prc.welcomeMessage = "Spreadsheet Magic @ Into The Box !!";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/index" );
	}


	function demo1( event, rc, prc ) {
		prc.welcomeMessage = "Demo 1 - Formulas";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/demo-1" );
	}

	function demo2( event, rc, prc ) {
		prc.welcomeMessage = "Demo 2 - Formatting";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/demo-2");
	}

	function demo3( event, rc, prc ) {
		prc.welcomeMessage = "Demo 3 - Conditional Formatting";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/demo-3" );
	}

	function demo4( event, rc, prc ) {
		prc.welcomeMessage = "Demo 4 - Dynamic Chart";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/demo-4" );
	}

	function demo5( event, rc, prc ) {
		prc.welcomeMessage = "Demo 5 - Importing Data";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/demo-5" );
	}

	function demo6( event, rc, prc ) {
		prc.welcomeMessage = "Demo 6 - Exporting Data";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/demo-6" );
	}

	function demo7( event, rc, prc ) {
		prc.welcomeMessage = "Demo 7 - Bonus";
		event.setLayout( "Default");
		event.setView( "spreadsheet-demos/demo-7" );
	}

}