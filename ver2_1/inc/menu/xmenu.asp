<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Outzource - TimeOut</title>
<link href='http://fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css'>
<link href="css/chronograph.less" rel="stylesheet/less" type="text/css" />
<script>
  less = {
    env: "development",
    async: false,
    fileAsync: false,
    poll: 1000,
    functions: {},
    dumpLineNumbers: "comments",
    relativeUrls: false,
    rootpath: ""
  };
</script>
<script src="../js/less.js" type="text/javascript"></script>
<script src="js/modernizr.js"></script>
</head>

<body>
    <div class="fixed-navbar-hoz">
        <!--<button class="showLeft" class="menu_icon"></button>-->
        <h1 class="menu_logo-left">
            <a href="../login.asp">TimeOut</a>
        </h1>
        <nav class="dropdown-right">
            <ul>
                <li><a href="#"><span class="glyph icon-flag"></span></a></li>
                <li><a href="#"><span class="glyph icon-printer"></span></a></li>
                <li>
                    <a href="#"><span class="glyph icon-user"></span><span class="account-name"><%=session("user") %>&nbsp;&nbsp;</span><b class="caret-down"></b></a>
                    <ul>
                        <li><a href="#">Hjælp</a></li>
                        <li><a href="#">Fjern support</a></li>
                        <li><a href="#">Indstillinger</a></li>
                        <li><a href="../sesaba.asp">Log ud</a></li>
                    </ul>
                </li>
            <ul>
        </nav>
    </div>
    <div class="fixed-navbar-vert">
        <nav>
            <ul>
                <li class="showLeft tooltip-right" tooltip-title="Timeregistrering"><a href="timereg_akt_2006.asp"><span class="glyph icon-stopwatch"></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Projekter"><a href="#"><span class="glyph icon-file-cabinet"></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Rapporter"><a href="#"><span class="glyph icon-chart-graph2"></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Kunder"><a href="#"><span class="glyph icon-building"></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Medarbejdere"><a href="#"><span class="glyph icon-user"></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Lager"><a href="#"><span class="glyph icon-barcode "></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Økonomi"><a href="#"><span class="glyph icon-piggy-bank"></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Administration"><a href="#"><span class="glyph icon-tools"></span></a></li>
                <li class="showLeft tooltip-right" tooltip-title="Help-Desk"><a href="#"><span class="glyph icon-help"></span></a></li>
            </ul>
        </nav>
    </div>
    <nav class="menu-slider menu-slider-left" id="menu-slider-1">
            <h3>Tidsregistrering</h3>
            <summary>I denne sektion har vi samlet alt vedr. Timeregistrering. Du har her mulighed for at registrere timer eller få et overblik over dine nuværende registreringer</summary>
            <ul>
                <li><a href="index.html">Timeregistrering</a></li>
                <li><a href="index.html">Ugeplan</a></li>
                <li><a href="index.html">Afstemning</a></li>
            </ul>   
        </nav>
    </body>
    </html>