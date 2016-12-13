

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/topmenu_inc.asp"-->

<!--#include file="../inc/regular/header_lysblaa_2015_inc.asp"-->
<!--#include file="../inc/regular/stat_func.asp"-->


<style type="text/css">

    tr
    {
        border-top:hidden;
    }



</style>


<div class="wrapper">
    <div class="content">


        <div class="container">
            <div class="porlet">
                <div class="portlet-body">

                    <div class="row">
                        <h3 class="col-lg-12">Kørsel</h3>
                    </div>


                    <table class="table">
                        <tr>
                            <td style="width:15%">Fra: <span style="color:red;">*</span>&nbsp</td>
                            <td><input type="text" class="form-control input-small" value="A. P. Møllers Allé 11, 2791 Dragø" /></td>                            
                        </tr>
                        <tr>
                            <td>Via: &nbsp</td>
                            <td><input type="text" class="form-control input-small" value="Energivej 230, 5020 Odense C" /></td>                            
                        </tr>
                        <tr>
                            <td>Via: &nbsp</td>
                            <td><input type="text" class="form-control input-small" value="Vestergade 2, 2871, KBH" /></td>                            
                        </tr>
                        <tr>
                            <td>Via: &nbsp</td>
                            <td><input type="text" class="form-control input-small" value="Nørregade 6, 2871, KBH" /></td>                            
                        </tr>
                        <tr>
                            <td>Til: <span style="color:red;">*</span>&nbsp</td>
                            <td><input type="text" class="form-control input-small" value="Nørregade 8, 2871, KBH" /></td>
                            <td style="border-top:hidden"><a href="#"><span style="color:darkred; display: block; text-align: center; padding-top:6px;" class="fa fa-times"></span></a></td>                            
                        </tr>
                        <tr style="border-top:hidden;">
                            <td><a href="#">Tilføj</a></td>
                        </tr>
                        
                    </table>

                    <br />

                    <table class="table">
                        <tr>
                            <td style="width:15%">Dato <span style="color:red;">*</span>&nbsp</td>
                            <td>
                                <div class='input-group date'>
                                      <input type="text" class="form-control input-small" name="FM_datoer" id="jq_dato" value="<%=day(now) &"/"& month(now) &"/"& year(now) %>" placeholder="dd-mm-yyyy" />
                                        <span class="input-group-addon input-small">
                                        <span class="fa fa-calendar">
                                        </span>
                                    </span>
                              </div>
                            </td>                            
                        </tr>
                        <tr>
                            <td>Kunde/Job: &nbsp</td>
                            <td><input type="text" class="form-control input-small" value="" /></td>                            
                        </tr>
                        <tr>
                            <td>Aktivitet: &nbsp</td>
                            <td><input type="text" class="form-control input-small" value="" /></td>                            
                        </tr>
                        <tr>
                            <td>Formål: <span style="color:red;">*</span>&nbsp</td>
                            <td><input type="text" class="form-control input-small" value="" /></td>
                            <td style="border-top:hidden"><a href="#"><span style="color:darkred; display: block; text-align: center; padding-top:6px; display:none" class="fa fa-times"></span></a></td>                            
                        </tr>
                        
                    </table>

                   

                  <!--  <table class="table">
                        <tr>
                            <td><span style="color:red;">*</span>&nbsp</td>
                            <td><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></td>                            
                        </tr>
                        <tr>
                            <td></td>
                            <td><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></td>                            
                        </tr>
                        <tr>
                            <td>Aktivitet: &nbsp</td>
                            <td><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></td>                            
                        </tr>

                        <tr style="border-top:hidden;">
                            <td>Formål: <span style="color:red;">*</span>&nbsp</td>
                            <td><input type="text" class="form-control input-small" value="" /></td>
                        </tr>
                        
                    </table> -->

                        <!--<div class="row">
                            <div class="col-lg-1">Fra:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">Via:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">Via:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">Via:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>
                        </div>
                        <div class="row">
                            <div class="col-lg-1">Fra:</div>
                            <div class="col-lg-4"><input type="text" class="form-control input-small" value="Hovgårdsvej 51, 7100 Vejle" /></div>
                        </div>

                        <div class="row">
                            <div class="col-lg-2"><a>Tilføj</a></div>
                        </div>

                      <!--  <div class="row">
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
                        </div> -->
                        
                        <div class="row">
                            <div class="col-lg-11"><button type="submit" class="btn btn-success btn-sm pull-right"><b>Indlæs >></b></button></div>
                        </div>
                    
                        <br /><br />

                </div>
            </div>
        </div>



    </div>
</div>

    

<!--#include file="../inc/regular/footer_inc.asp"-->