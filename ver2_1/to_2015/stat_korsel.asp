

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->

<%call menu_2014 %>

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
                            <div class="col-lg-1"><b>Dato:</b></div>
                            <div class="col-lg-3">
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
                            <div class="col-lg-1"><b>Kunde/Job:</b></div>
                            <div class="col-lg-3"><input type="text" class="form-control input-small" /></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1"><b>Aktivitet:</b></div>
                            <div class="col-lg-3"><input type="text" class="form-control input-small" /></div>
                        </div>
                        <br />
                        <div class="row">
                            <div class="col-lg-1"><span style="display: block; text-align:right; padding-top:5px" class="fa fa-circle-o"></span></div>
                            <div class="col-lg-3"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1"><span style="display: block; text-align:right; padding-top:5px" class="fa fa-circle-o"></span></div>
                            <div class="col-lg-3"><input type="text" class="form-control input-small" value="Dragørvej 45, 819 Dragør" /></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1"><span style="display: block; text-align:right; padding-top:5px" class="fa fa-child"></span></div>
                            <div class="col-lg-3"><input type="text" class="form-control input-small" value="Køkebnahvnvej 14, 2100, Købehavnvej" /></div>
                            <div class="col-lg-1"><a href="#"><span style="display: block; text-align:right; padding-top:5px" class="fa fa-times"></span></a></div>
                            <div class="col-lg-1"><a href="#" style="text-decoration:none;"><span style="display: block; text-align:left; padding-top:5px" class="fa fa-plus-circle"></span></a></div>
                        </div>

                        <br /><br /><br />
                        <div class="row">
                            <div class="col-lg-2"><button type="submit" class="btn btn-success btn-sm"><b>Indlæs >></b></button></div>
                        </div>
                    </div>

                    <br /><br />
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
                    </div>

                    <div class="row">
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
                    </div>

                    <br />
                    <div class="row">
                        <div class="col-lg-2"><button type="submit" class="btn btn-secondary btn-sm"><b>Vis >></b></button></div>
                    </div>

                    <br /><br />

                    <table class="table">
                        <thead>
                            <tr>
                                <th style="width:10%">Dato</th>
                                <th style="width:20%">Fra</th>
                                <th style="width:20%">Til</th>
                                <th style="width:10%">Kommentar</th>
                                <th style="width:10%">Projekt</th>
                                <th style="width:5%">Retu</th>
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
                                <td>Olivers liv</td>
                                <td class="center"><input type="checkbox" value="1" /></td>
                                <td><input type="text" class="form-control input-small" value="245" /></td>
                                <td><input type="text" class="form-control input-small" value="500" /></td>
                                <th><input type="text" class="form-control input-small" value="500" /></th>
                            </tr>
                        </tbody>
                    </table>

                </div>
            </div>
        </div>
    </div>
</div>
    

<!--#include file="../inc/regular/footer_inc.asp"-->