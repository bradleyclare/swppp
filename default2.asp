<%
base_path = server.mappath(".")

' iomode settings
ForReading = 1
ForWriting = 2
ForAppending = 8

'format settings
TristateUseDefault = -2
TristateTrue = -1
TristateFalse = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")
'response.write(base_path & "/about.txt")
Set objFile = objFSO.OpenTextFile(base_path & "/static/about.txt")
%>

<html>
<head>
	<title>SWPPP INSPECTIONS</title>
	<link rel="stylesheet" type="text/css" href="global.css">
</head>
<title>SWPPP INSPECTIONS : Home Page</title>
    <link rel="stylesheet" type="text/css" href="css/bootstrap.min.css" />
	<link rel="stylesheet" type="text/css" href="css/carousel.css" />
	<link rel="stylesheet" type="text/css" href="css/my_bootstrap.css" />
	<script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <link rel="stylesheet" type="text/css" href="global.css">
</head>
<body bgcolor="#FFFFFF" text="#000000">
    <div class="navbar-wrapper">
        <div class="container">
            <div class="navbar navbar-inverse navbar-static-top" role="navigation">
                <div class="container">
                    <div class="navbar-header">
                        <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target=".navbar-collapse">
                            <span class="sr-only">Toggle navigation</span>
                            <span class="icon-bar"></span>
                            <span class="icon-bar"></span>
                            <span class="icon-bar"></span>
                        </button>
                        <a class="navbar-brand" href="#">
                            <div class="mobile-only">
                                <img src="./images/logo_top_white_small.png" /></div>
                            <div class="desktop-only">
                                <img src="./images/logo_top_white.png" /></div>
                        </a>
                    </div>
                    <div class="navbar-collapse collapse">
                        <ul class="nav navbar-nav">
                            <li class="active"><a href="index">Home</a></li>
                            <li><a href="views/projects.asp">Log In</a></li>
                            <li><a href="admin/default.asp">Admin</a></li>
                        </ul>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <div class='well'>
        <h1>SWPPP INSPECTIONS, INC.</h1>
    </div>
    <div class='well'>
        <div class='row featurette'>
            <div class='col-md-5'>
                <h3>WHO WE ARE:</h3>
                <p>We specialize in compliance with the Texas Commission on Environmental Quality's 
                    (TCEQ) Permit TXR150000 and the Oklahoma Department of Environmental Quality's 
                    (ODEQ) Permit OKR10, which govern stormwater runoff from construction sites. We 
                    serve departments of transportation, municipalities, civil engineers, developers, 
                    homebuilders, general contractors, and other construction related subcontractors 
                    (including dirt, utility, paving, and erosion and sediment control companies).</p>
            </div>
            <div class='col-md-5'>
                <h3>SERVICES:</h3>
                <ul>
                    <li>Storm Water Pollution Prevention Plan design or SWPPP, SWP3, SW3P design</li>
                    <li>Weekly emailed inspection reports</li>
                    <li>Turnkey compliance</li>
                    <li>Pre-construction consultation and training with city,    
                     project engineers, and all project subcontractors</li>
                    <li>Online and onsite documentation management</li>
                    <li>Information management: password-protected online access to the SWPPP, related 
                        documents, inspections, Site Maps, and site photographs</li>
                </ul>
            </div>
        </div>
    </div>
    <div class='well'>
        <h2>To request a proposal or receive more information, contact us today.</h2>
        <center><strong>
        SWPPP INSPECTIONS, INC.<br>
        PO Box 496987<br>
        Garland, TX 75049<br>
        972.530.5307 Office<br>
        972.530.5309 Fax<br>
        <a href="mailto:info@swppp.com">info@swppp.com</font></a>
        </strong></center>
    </div>
    <div class='well'>
        <div class='row featurette'>
            <h3>SERVICE AREAS:</h3>
            <strong>North Texas counties:</strong> Bosque County, Collin County, Cooke County, Dallas County, Delta County, Denton County, Ellis County, Grayson County, Henderson County, Hill County, Hunt County, Johnson County, Kaufman County, McLennan County, Navarro County, Parker County, Rockwall County, Tarrant County, Van Zandt County, and Wise County<br/><br/>

            <strong>North Texas cities:</strong> Addison, Aledo, Allen, Alma, Alvarado, Anna, Annetta, Antioch, Argyle, Arlington, Athens, Aubrey, Azle, Balch Springs, Bardwell, Bartonville, Baxter, Bedford, Bells, Benbrook, Berryville, Blue Mound, Blue Ridge, Briar, Briaroaks, Brownsboro, Burleson, Caddo Mills, Campbell, Caney City, Carrollton, Cedar Hill, Celeste, Celina, Chandler, Cleburne, Cockrell Hill, Coffee City, Colleyville, Collinsville, Combine, Commerce, Cool, Coppell, Copper Canyon, Corinth, Corral City, Corsicana, Cottonwood, Crandall, Cross Roads, Cross Timber, Crowley, DFW, Dallas, Dalworthington Gardens, Decatur, Denison, Denton, DeSoto, Dorchester, Double Oak, Duncanville, Eagle Mountain, Edgecliff Village, Enchanted Oaks, Ennis, Euless, Eustace, Everman, Fairview, Farmers Branch, Farmersville, Fate, Ferris, Flower Mound, Forest Hill, Forney, Fort Worth, Frisco, Garland, Garrett, Glenn Heights, Godley, Gordonville, Grand Prairie, Grandview, Grapevine, Grays Prairie, Greenville, Gun Barrel City, Gunter, Hackberry, Haltom City, Haslet, Hawk Cove, Heath, Hebron, Hickory Creek, Highland Park, Highland Village, Hillsboro, Howe, Hudson Oaks, Hurst, Hutchins, Irving, Italy, Josephine, Joshua, Justin, Kaufman, Keene, Keller, Kemp, Kennedale, Knollwood, Krugerville, Krum, Lake Dallas, Lake Worth, Lakeside, Lakewood Village, Lancaster, Larue Lavon, Leagueville, Lewisville, Lincoln Park, Lipan, Little Elm, Log Cabin, Lone Oak, Lowry Crossing, Lucas, Mabank, Malakoff, Mansfield, Marshall Creek, Maypearl, McKinney, McLendon-Chisholm, Melissa, Meridian, Mesquite, Midlothian, Midway, Milford, Millsap, Mobile City, Moore Station, Murchison, Murphy, Nevada, New Hope, New York, Newark, Neylandville, North Richland Hills, Northlake Oak Cliff, Oak Grove, Oak Leaf, Oak Point, Oak Ridge, Oak Trail Shores, Opelika, Ovilla, Palmer, Pantego, Parker, Payne Springs, Pecan Acres, Pecan Hill, Pecan Plantation, Pelican Bay, Pilot Point, Plano, Ponder, Post Oak Bend City, Pottsboro, Poynor, Princeton, Prosper, Quinlan, Red Branch, Red Oak, Rendon, Reno, Richardson, Richland Hills, Rio Vista, River Oaks, Roanoke, Rockwall, Rosser, Rowlett, Royse City, Sachse, Sadler, Saginaw, Saint Paul, Sanctuary, Sanger, Sansom Park, Seagoville, Seven Points, Shady Shores, Sherman, Southlake, Southmayd, Springtown, Star Harbor, Sunnyvale, Talty. Telico, Terrell, The Colony, Tioga, Tolar, Tom Bean, Tool, Trinidad, Trophy Club, University Park, Van Alstyne, Venus Watauga, Waco, Waxahachie, Weatherford, West Tawakoni, Westlake, Westminster, Weston, Westover Hills, Westworth Village, White Settlement, Willow Park, Wilmer, Wolfe City, and Wylie<br/><br/>

            <strong>West Texas counties:</strong> Midland County and Ector County<br/><br/>

            <strong>West Texas cities:</strong> Midland and Odessa<br/><br/>

            <strong>Oklahoma counties:</strong> Canadian County, Cleveland County, Grady County, Lincoln County, Logan County, McClain County, and Oklahoma County<br/><br/>

            <strong>Oklahoma cities:</strong> Bethany, Bethel Acres, Blanchard, Bridge Creek, Chandler, Chickasha, Choctaw, Del City, Edmond, El Reno, Goldsby, Guthrie, Harrah, Jones, Kingfisher, Lexington, McLoud, Meeker, Midwest City, Minco, Moore, Mustang, Newcastle, Nichols Hills, Nicoma Park, Noble, Norman, Okarche, Piedmont, Pink, Purcell, Shawnee, Slaughterville, Spencer, Tecumseh, The Village, Turtle, Union City, Valley Brook, Warr Acres, Washington, and Yukon
        </div>
    </div>
</body>
</html>