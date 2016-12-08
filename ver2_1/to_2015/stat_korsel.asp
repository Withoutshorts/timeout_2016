

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->






<style type="text/css">

    .center 
    {
        text-align:center;
    }

</style>



<div class="wprapper">
    <div class="content">

        <div class="container">
            <div class="portlet">
                <h3 class="portlet-title"><u>Kørsel</u></h3>
                <div class="portlet-body">
                
                    <div class="well well-white">

                        <div class="row">
                            <div class="col-lg-1">Fra:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>

                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-1">Dato:</div>
                            <div class="col-lg-4">
                                <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            </div>
                        </div>
                        
                        <div class="row">
                            <div class="col-lg-1">Via:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Dragørvej 45, 819 Dragør"/></div>

                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-1">Kunde/Job:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" /></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-1">Til:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Køkebnahvnvej 14, 2100, Købehavnvej" /></div>
                            <div class="col-lg-1">&nbsp</div>
                            <div class="col-lg-1">Aktivitet:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" /></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-2"><a>Tilføj</a></div>
                        </div>
                        
                        <div class="row">
                            <div class="col-lg-11"><button type="submit" class="btn btn-success btn-sm pull-right"><b>Indlæs >></b></button></div>
                        </div>
                    </div>


                    <br /><br /><br /><br />


                    <div class="row">
                        <h5 class="col-lg-6">Kørsels statistik</h5>
                    </div>
                    <div class="row">
                        <div class="col-lg-1">Medarbejder:</div>
                        <div class="col-lg-4">
                            <select class="form-control input-small">
                                <option value="1">Jesper Jespersen Aaberg</option>
                                <option value="1">Søren</option>
                                <option value="1">Storm</option>
                            </select>
                        </div>
                        
                        <div class="col-lg-1">&nbsp</div>

                        <div class="col-lg-1">Periode:</div>
                        <div class="col-lg-2">
                            <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                            </div>  
                        </div>
                        <div class="col-lg-2">
                            <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                            </div>  
                        </div>   
                        <div class="col-lg-1"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis >></b></button></div>                     
                    </div>


                    <br />
                   <!-- <div class="row">
                        <div class="col-lg-11"><button type="submit" class="btn btn-secondary btn-sm pull-right"><b>Vis >></b></button></div>
                    </div> -->

                    <br /><br />

                    <table class="table">
                        <thead>
                            <tr>
                                <th style="width:10%">Dato</th>
                                <th style="width:20%">Fra</th>
                                <th style="width:20%">Til</th>
                                <th style="width:10%">Formål</th>
                                <th style="width:10%">Projekt</th>
                                <th style="width:5%">Retur</th>
                                <th style="width:5%">Kørte km</th>
                                <th style="width:5%">km i alt</th>
                                <th style="width:5%">Afregnet</th>
                            </tr>
                        </thead>

                        <tbody>
                            <tr>
                                <td>08-12-2016</td>
                                <td>Vejle</td>
                                <td>Kbh</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="250" /></td>
                                <td><input type="text" class="form-control input-small" value="500" /></td>
                                <th class="center"><input type="checkbox"/></th>
                            </tr>
                            <tr>
                                <td>12-12-2016</td>
                                <td>A. P. Møllers Allé 11, 2791 Dragø</td>
                                <td>Energivej 230, 5020 Odense C</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="150" /></td>
                                <td></td>
                                <th></th>
                            </tr>
                            <tr>
                                <td></td>
                                <td>Energivej 230, 5020 Odense C</td>
                                <td>En gade vest på, 3333 ogenby</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="100" /></td>
                                <td></td>
                                <th></th>
                            </tr>
                            <tr>
                                <td></td>
                                <td>Vestergade 2, 2871, KBH</td>
                                <td>Nørregade 6, 2871, KBH</td>
                                <td>Møde</td>
                                <td>TimeOut Opsætning & Support (527)</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="20" /></td>
                                <td><input type="text" class="form-control input-small" value="270" /></td>
                                <th class="center"><input type="checkbox"/></th>
                            </tr>
                            
                        </tbody>
                    </table>

                    <a data-toggle="modal" href="#korsel_pop"><span class="fa fa-info-circle"></span></a>

                    <div id="korsel_pop" class="modal modal-styled fade" style="top:60px;"><!-- modal modal-styled fade -->
                        <div class="modal-dialog">
                            <div class="modal-content" style="border:none !important;padding:0;">
                                <div class="modal-header">
                                <button type="button" class="close" data-dismiss="modal" aria-hidden="true">&times;</button>
                                    <h5 class="modal-title">Medarbejdertype</h5>
                                </div>
                                <div class="modal-body">

                                    hej

                                </div>
                            </div>
                        </div>
                    </div>



                </div>
            </div>
        </div>
    </div>
</div>
    

<!--#include file="../inc/regular/footer_inc.asp"-->