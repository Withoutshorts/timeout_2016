<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>

    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

    <!-- Googlefonrt apis -->
    <link href='//fonts.googleapis.com/css?family=Open+Sans:400,300,600,700,800' rel='stylesheet' type='text/css' />
     
    <!-- Bruges til LESS styling af menu -->
    <link href="css/chronograph.css" rel="stylesheet" type="text/css" />

    <!-- Bootstrap core CSS and indivisual styles for TimeOut -->
    <link href="assets/css/bootstrap.css" rel="stylesheet" />

    <!-- Custom styles for DASH GUM -->
    <link href="assets/css/style.css" rel="stylesheet" />
    <link href="assets/css/style-responsive.css" rel="stylesheet" />
    <link href="assets/css/table-responsive.css" rel="stylesheet" /> 
    <link href="assets/font-awesome/css/font-awesome.css" rel="stylesheet" />

    <!-- Custom styles for TimeOut - DASHGUM CSS folder  -->
    <link href="assets/css/style-timeout.css" rel="stylesheet" />

    <!-- Include the jQuery library  -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js"></script>

    <!-- Include the Bootstrap JavaScript plugins  -->
    <script src="http://maxcdn.bootstrapcdn.com/bootstrap/3.2.0/js/bootstrap.min.js"></script>

    <!-- aTOOLTIP SCRIPT -->    
    <script type="text/javascript" src="aToolTip/js/jquery.atooltip.min.js"></script>
    
    <!-- aTOOLTIP - TEXTS FOR TIMEOUT --> 
    <script type="text/javascript" src="atooltip/js/timeout_tooltips.js"></script>

	<!-- aTOOLTIP CSS -->
	<link type="text/css" href="aToolTip/css/atooltip.css" rel="stylesheet"  media="screen" />

    <!--#include file="xmenu.asp"-->

</head>
<body>

<div class="main-marg-timeout">
    <div class="row mt">
        <div class="col-lg-12">
            <div class="content-panel">
                <!-- HEADER FOR TABEL -->
                <section>
                    <div class="row head-marg-tbl-col12">
                        <div class="col-lg-11"><h4><i class="fa fa-angle-right"></i> Medarbejdertyper</h4></div>
                        <div class="col-lg-1"><button type="button" class="btn btn-success btn-sm">&nbsp;&nbsp;&nbsp;<b>Opret</b>&nbsp;&nbsp;&nbsp;</button></div>
                    </div>
                </section>
                <section>
                    <table class="table table-bordered table-striped table-condensed">
                        <thead>
                            <tr class="bg-primary">
                                <th>Id</th>
                                <th>Sortering</th>
                                <th class="numeric">Navn</th>
                                <th class="numeric">Hovedgruppe</th>
                                <th class="numeric">Søstergruppe</th>                              
                                <th class="numeric"><a href="#" class="clickTipMember" style="color:white;">Medlemmer&nbsp;&nbsp;&nbsp;<span class="glyphicon glyphicon-info-sign"></span></a></th>
                                <th class="numeric">Gnr. timepris</th>
                                <th class="numeric">Intern kostpris</th>
                                <th class="numeric"><a href="#" class="clickTipHW" style="color:white;">Timer/uge&nbsp;&nbsp;&nbsp;<span class="glyphicon glyphicon-info-sign"></a></th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td>4</td>
                                <td>0</td>
                                <td class="numeric"><a href="">Designer</a></td>
                                <td class="numeric">Medarbejdertype Hove</td>
                                <td class="numeric">(id: 15)</td>
                                <td class="numeric"><a href="">medlemmer</a> (4) (1)</td>
                                <td class="numeric">800 DKK</td>
                                <td class="numeric">450 DKK</td>
                                <td class="numeric">37,00 timer (7,4)</td>
                            </tr>
                            <tr>
                                <td>1</td>
                                <td>0</td>
                                <td class="numeric"><a href="">Direktør</a></td>
                                <td class="numeric">Medarbejdertype Hove</td>
                                <td class="numeric">(id: 15)</td>
                                <td class="numeric"><a href="">medlemmer</a> (4) (1)</td>
                                <td class="numeric">650 DKK</td>
                                <td class="numeric">250 DKK</td>
                                <td class="numeric">37,00 timer (7,4)</td>
                            </tr>
                            <tr>
                                <td>14</td>
                                <td>0</td>
                                <td class="numeric"><a href="">Marketing</a></td>
                                <td class="numeric">Medarbejdertype Hove</td>
                                <td class="numeric">(id: 15)</td>
                                <td class="numeric"><a href="">medlemmer</a> (4) (1)</td>
                                <td class="numeric">800 DKK</td>
                                <td class="numeric">0 DKK</td>
                                <td class="numeric">37,00 timer (7,4)</td>
                            </tr>
                            <tr>
                                <td>2</td>
                                <td>0</td>
                                <td class="numeric"><a href="">Projektleder</a></td>
                                <td class="numeric">Medarbejdertype Hove</td>
                                <td class="numeric">(id: 15)</td>
                                <td class="numeric"><a href="">medlemmer</a> (4) (1)</td>
                                <td class="numeric">650 DKK</td>
                                <td class="numeric">0 DKK</td>
                                <td class="numeric">37,00 timer (7,4)</td>
                            </tr>
                            <tr>
                                <td>15</td>
                                <td>0</td>
                                <td class="numeric"><a href="">Sælger</a></td>
                                <td class="numeric">Medarbejdertype Hove</td>
                                <td class="numeric">(id: 15)</td>
                                <td class="numeric"><a href="">medlemmer</a> (4) (1)</td>
                                <td class="numeric">800 DKK</td>
                                <td class="numeric">350 DKK</td>
                                <td class="numeric">37,00 timer (7,4)</td>
                            </tr>
                            <tr>
                                <td>8</td>
                                <td>0</td>
                                <td class="numeric"><a href="">Udvikler</a></td>
                                <td class="numeric">Medarbejdertype Hove</td>
                                <td class="numeric">(id: 15)</td>
                                <td class="numeric"><a href="">medlemmer</a> (4) (1)</td>
                                <td class="numeric">650 DKK</td>
                                <td class="numeric">350 DKK</td>
                                <td class="numeric">37,00 timer (7,4)</td>
                            </tr>
                        </tbody>
                    </table>
                </section>
            </div>
        </div>
    </div>

    <div class="row mt">
        <div class="col-lg-10">
            <div class="content-panel">
                <!-- EDITÈRBAR HEADER -->
                <section>
                    <div class="row head-marg-form">
                        <div class="col-lg-10"><h4><i class="fa fa-angle-right"></i> Redigér medarbejdertyper</h4></div>                            
                        <div class="col-lg-1"><a href="javascript.history.back()"><button type="button" class="btn btn-danger btn-sm" /><b>Annullér</b></button></a></div>
                        <div class="col-lg-1"><button type="button" class="btn btn-primary btn-sm"><b>Opdatér</b></button></div>
                    </div>
                </section>
                <!-- NAVN / SORTERING / ID -->
                <section class="sec-marg-gen">
                    <div class="row">
                        <!-- Navn -->
                        <div class="col-lg-6 col-padR-5">
                            <div class="well">
                                <div class="row">
                                    <div class="col-lg-4 col-padL-10"><label>Medarbejdernavn:</label></div>
                                    <div class="col-lg-8">   
                                        <input type="text" class="form-control input-sm" value="" />
                                    </div>
                                </div>
                            </div>
                        </div>
                        <!-- Sortering -->
                        <div class="col-lg-3 col-padR-5">
                            <div class="well">
                                <div class="row">
                                    <div class="col-lg-6 col-padL-10"><label>Sortering:</label></div>
                                    <div class="col-lg-6">   
                                        <input type="text" class="form-control input-sm" />
                                    </div>
                                </div>
                            </div>
                        </div>
                        <!-- Id -->
                        <div class="col-lg-3">
                            <div class="well">
                                <div class="row">
                                    <div class="col-lg-6 col-padL-10"><label>Id:</label></div>
                                    <div class="col-lg-6">   
                                        <input type="text" class="form-control input-sm text-center full-width" value="12" disabled="disabled"" />
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </section>
                <!-- NORMERET TID -->
                <section class="sec-marg-gen">
                    <div class="well">
                        <div class="row">
                            <div class="col-lg-2 col-padL-10">
                                <label>Normeret tid&nbsp;&nbsp;&nbsp;<a href="#" class="clickTipNrmTidFerie">
                                    <span class="glyphicon glyphicon-info-sign"></span></a>
                                </label>
                            </div>
                            <div class="col-lg-10 col-padL-25">   
                                <div class="row">
                                    <div class="col-sm-1 text-center col-pad-week">
                                        <div class="text-center">Mandag</div>
                                        <input type="text" class="form-control input-sm" />
                                    </div>
                                    <div class="col-sm-1 text-center col-pad-week">
                                        <div class="text-center">Tirsdag</div>
                                        <input type="text" class="form-control input-sm" />
                                    </div>
                                    <div class="col-sm-1 text-center col-pad-week">
                                        <div class="text-center">Onsdag</div>
                                        <input type="text" class="form-control input-sm" />
                                    </div>
                                    <div class="col-sm-1 text-center col-pad-week">
                                        <div class="text-center">Torsdag</div>
                                        <input type="text" class="form-control input-sm" />
                                    </div>
                                    <div class="col-sm-1 text-center col-pad-week">
                                        <div class="text-center">Fredag</div>
                                        <input type="text" class="form-control input-sm" />
                                    </div>
                                    <div class="col-sm-1 text-center col-pad-week">
                                        <div class="text-center">Lørdag</div>
                                        <input type="text" class="form-control input-sm input-red" />
                                    </div>
                                    <div class="col-sm-1 text-center col-pad-week col-pad-week-son">
                                        <div class="text-center">Søndag</div>
                                        <input type="text" class="form-control input-sm input-red" />
                                    </div>
                                    <div class="col-sm-1 text-center col-pad-week" style="margin-right:0;">
                                        <div class="text-center"><b>Timetotal</b></div>
                                        <input type="text" disabled="disabled" class="form-control input-sm" placeholder="84,00" />
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </section>
                <!-- Timepris generel -->
                <section class="sec-marg-gen">
                    <div class="well">
                        <div class="row">
                            <div class="col-lg-2 col-padL-10">
                                <label>Timepris generel&nbsp;&nbsp;&nbsp;<a href="#" class="clickTipTmPrsGn">
                                    <span class="glyphicon glyphicon-info-sign"></span></a>
                                </label>
                            </div>
                            <div class="col-sm-2" style="margin-left:-5px;">
                                <input type="text" class="form-control input-sm" value="300" />
                            </div>
                            <div class="dropdown col-sm-2">
                                <select class="form-control input-sm">
						            <option>DKK</option>
						            <option>EUR (7,4556)</option>
						            <option>SEK (0,8934)</option>
					            </select>
                            </div>
                            <div class="col-sm-6" style="padding-left:82px;">                                            
                                <!-- Åbn/luk Flere timepriser -->
                                <div style="margin-top:10px;" class="text-left">
                                    <a href=".details" data-toggle="collapse" style="cursor: pointer;">
                                        <div class="details collapse in">
                                            <div class="TmPrFloatLtxt">Se flere timepriser&nbsp;&nbsp;</div>
                                            <div class="TmPrFloatLtxtGlyph" style="margin-top:-2px;"><span class="glyph icon-arrow-down"></span></div>
                                        </div>
                                        <div class="details collapse">
                                            <div class="TmPrFloatLtxt">Luk flere timepriser&nbsp;&nbsp;</div>
                                            <div class="TmPrFloatLtxtGlyph" style="margin-top:-5px;"><span class="glyph icon-arrow-up"></span></div>
                                        </div>
                                        <div style="clear:both;"></div>
                                    </a>
                                </div>
                            </div>
                        </div>
                        <div class="details collapse" style="margin-left:-7px;">
                            <!-- Timepris 2+3 -->
                            <div class="row mt">
                                <label class="col-sm-2 text-right col-pad-tmpr">Timepris 2:</label>
                                <div class="col-sm-2"><input type="text" class="form-control input-sm" /></div>
                                <div class="dropdown col-sm-2">
                                    <select class="form-control input-sm">
						                <option>DKK</option>
						                <option>EUR (7,4556)</option>
						                <option>SEK (0,8934)</option>
					                </select>                                            
                                </div>
                                <label class="col-sm-2 text-right col-pad-tmpr">Timepris 3:</label>
                                <div class="col-sm-2"><input type="text" class="form-control input-sm" /></div>
                                <div class="dropdown col-sm-2">
                                    <select class="form-control input-sm">
						                <option>DKK</option>
						                <option>EUR (7,4556)</option>
						                <option>SEK (0,8934)</option>
					                </select>                                            
                                </div>
                            </div>
                            <!-- Timepris 4+5 -->
                            <div class="row mt">
                                <label class="col-sm-2 text-right col-pad-tmpr">Timepris 4:</label>
                                <div class="col-sm-2"><input type="text" class="form-control input-sm" /></div>
                                <div class="dropdown col-sm-2">
                                    <select class="form-control input-sm">
						                <option>DKK</option>
						                <option>EUR (7,4556)</option>
						                <option>SEK (0,8934)</option>
					                </select>                                            
                                </div>
                                <label class="col-sm-2 text-right col-pad-tmpr">Timepris 5:</label>
                                <div class="col-sm-2"><input type="text" class="form-control input-sm" /></div>
                                <div class="dropdown col-sm-2">
                                    <select class="form-control input-sm">
						                <option>DKK</option>
						                <option>EUR (7,4556)</option>
						                <option>SEK (0,8934)</option>
					                </select>                                            
                                </div>
                            </div>
                        </div>
                    </div>
                </section>
                <!-- Intern kostpris -->
                <section class="sec-marg-gen">
                    <div class="well">
                        <div class="row">
                            <div class="col-lg-2 col-padL-10"><label>Intern kostpris</label></div>
                            <div class="col-sm-2" style="margin-left:-5px;">
                                <input type="text" class="form-control input-sm" />
                            </div>
                            <div class="col-sm-1" style="padding:8px 0 0 0">DKK</div>
                            <div class="col-sm-7" style="margin-top:-3px;">
                                <div class="checkbox">
                                    <label>
                                        <input type="checkbox" value="" />
                                        Opdater interne kostpriser, for denne type, på alle eksisterende <u>åbne</u> job.
                                    </label>
                                </div>
                            </div>
                        </div>
                    </div>
                </section>
                <!-- SØSTER/HOVEDGRUPPER -->
                <section class="sec-marg-gen">
                    <div class="row">
                        <div class="col-lg-6 col-padR-5">
                            <div class="well">
                                <div class="row">
                                    <label class="col-sm-4 col-padL-10">
                                        Søster-gruppe&nbsp;&nbsp;&nbsp;<a href="#" class="clickTipSoestGrp"><span class="glyphicon glyphicon-info-sign"></span></a> 
                                    </label>
                                    <div class="col-sm-8">
                                        <select class="form-control input-sm">
						                    <option>Nej</option>
						                    <option>Designer</option>
						                    <option>Marketing</option>
					                    </select>
                                    </div>
                                </div>
                            </div>
                        </div>
                        <div class="col-lg-6">
                            <div class="well">
                                <div class="row">
                                    <label class="col-sm-4 col-padL-10">
                                        Hoved-gruppe&nbsp;&nbsp;&nbsp;<a href="#" class="clickTipHvdGrp"><span class="glyphicon glyphicon-info-sign"></span></a>
                                    </label>
                                    <div class="col-sm-8">
                                        <select class="form-control input-sm">
						                    <option>Medarbejdertype Hovedgruppe 1</option>
						                    <option>//Følger søstergruppe//</option>
					                    </select>
                                    </div>
                                </div>
                                </div>
                        </div>
                    </div>
                </section>
                <!-- AFSLUT UGE KRITERIE -->
                <section class="sec-marg-gen">
                    <div class="well">
                        <div class="row">
                            <label class="col-sm-2 col-padL-10">Afslut uge kriterie</label>
                            <div class="col-sm-3" style="padding:0 0 0 10px;">
                                <div>Minimumskrav til timeantal.</div>
                                <div class="small">Medarbejdertypen kan afslutte sin uge når:</div>
                            </div>
                            <div class="col-sm-3" style="padding:0 10px 0 0;">
                                <select class="form-control input-sm">
						            <option>Ingen (kan altid afslutte)</option>
						            <option>Realiserede timer</option>
                                    <option>Realiserede fakturérbare timer</option>
					            </select>
                            </div>
                            <div class="col-sm-1" style="padding:5px 0 0 0;">udgør min.:</div>
                            <div class="col-sm-2" style="padding:0 5px 0 0;"><input type="text" class="form-control input-sm" /></div>
                            <div class="col-sm-1" style="padding:5px 0 0 0;">% af normtid.</div>
                        </div>
                    </div>
                </section>
                <!-- OPDATÈR STAMAKTIVITETSGRUPPER -->
                <section class="sec-marg-gen">
                    <div class="well">
                        <div class="form-group">
                            <!-- <div class="detailsAktGrp collapse in">
                                <div class="row">
                                    <label class="col-sm-2 control-label">
                                        <a data-toggle="collapse" href=".detailsAktGrp">
                                            Åbn Opdatér stam-aktivitets-grp.&nbsp;&nbsp;<span class="glyph icon-arrow-down"></span>
                                        </a>
                                    </label>
                                </div>
                            </div>
                                -->
                            <!-- <div class="detailsAktGrp collapse"> -->
                                <div class="row">
                                    <label class="col-sm-2 col-padL-10">
                                        Opdatér stam-aktivitets-grp.
                                        <!--
                                        <a data-toggle="collapse" href=".detailsAktGrp">
                                            <span>Luk Opdatér stam-aktivitets-grp.</span>&nbsp;&nbsp;<span class="glyph icon-arrow-up"></span>
                                        </a>
                                        -->
                                    </label>
                                    <div class="col-sm-4">
                                        <select multiple class="form-control input-sm">
                                            <option>Standard 2014</option>
                                            <option>Timeout</option>
                                            <option>TO implementering</option>
                                            <option>Udviklingsprojekt</option>
                                        </select>
                                    </div>
                                    <div class="col-sm-6">
                                        <div class="checkbox">
                                            <label style="width:90%;margin-top:-10px;">
                                                <input type="checkbox" value="" />
                                                Ja, opdatér timepriserne for alle medarbejdere af denne type på alle aktiviteterne, i stamaktivitets-grupperne til venstre:
                                            </label>
                                        </div>
                                    </div>
                                </div>
                            <!-- </div> -->
                        </div>
                    </div>
                </section>
            </div>
        </div>
    </div>

</div>

<br/><br/><br/><br/>

</body>
</html>
