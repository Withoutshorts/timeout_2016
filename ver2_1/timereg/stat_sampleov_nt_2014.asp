

<%response.buffer = true%>

<!--#include file="../inc/connection/conn_db_inc.asp"-->
<!--#include file="../inc/errors/error_inc.asp"-->
<!--#include file="../inc/regular/global_func.asp"-->
<!--#include file="../inc/regular/job_func.asp"-->
<!--#include file = "CuteEditor_Files/include_CuteEditor.asp" -->


<!--#include file="../inc/regular/header_lysblaa_2014_inc.asp"--> 
<!--#include file="../inc/regular/topmenu_inc.asp"--> 

<%call menu_2014 %>

<div class="container">

   
        <div class="row">
            <!--SEARCH START -->
            <div id="search" class="col-s-12 col-m-6">
                <section class="panel">
                    <form class="panel-body">
                        <div class="search-group">
                        <input type="search"/>
                        <span class="search-group-btn"><button type="button">Søg</button></span>
                        </div>
                        <div class="col-s-12 no-pad-hoz">
                            <label class="checkbox-inline"><input type="checkbox" vakue="Enquiry">Enquiry</label>
                            <label class="checkbox-inline"><input type="checkbox" vakue="Active orders">Active orders</label>
                            <label class="checkbox-inline"><input type="checkbox" vakue="Shipped orders">Shipped orders</label>
                            <label class="checkbox-inline"><input type="checkbox" vakue="Closed orders">Closed orders</label>
                        </div>
                    </form>
                </section>
            </div>
            <!--SEARCH END -->
            <!--TABLE START -->
            <div id="search" class="col-s-12">
                <section class="panel">
                    <header class="panel-heading">table auto</header>
                    <table class="table table-auto">
                        <thead>
                            <tr>
                                <th data-title="indberettet på aktivitet og heraf fakturerbart">Style</th>
                                <th>Buyer</th>
                                <th>Supplier</th>
                                <th>Src DL</th>
                                <th>Proto DL</th>
                                <th>Photo DL</th>
                                <th>SMS DL</th>
                                <th>Cost P</th>
                                <th>Sales P</th>
                                <th>Profit/PC</th>
                            </tr>
                        </thead>
                        <tbody>
                            <tr>
                                <td><a href="#">SS #15</a></td>
                                <td><a href="#">Just-Female</a></td>
                                <td><a href="#">Supplier</a></td>
                                <td><a href="#">10-05</a></td>
                                <td><a href="#">15-05</a></td>
                                <td><a href="#">20-05</a></td>
                                <td><a href="#">20-05</a></td>
                                <td><a href="#">1.000.000</a></td>
                                <td><a href="#">1.505.000</a></td>
                                <td><a href="#">1.000.000</a></td>
                            </tr>
                            <tr>
                                <td>SS #15</td>
                                <td>Just-Female</td>
                                <td>Supplier</td>
                                <td>10-05</td>
                                <td>15-05</td>
                                <td>20-05</td>
                                <td>20-05</td>
                                <td>1.000.000</td>
                                <td>1.55.000</td>
                                <td>1.0000.000</td>
                            </tr>
                            <tr>
                                <td>SS #15</td>
                                <td>Just-Female</td>
                                <td>Supplier</td>
                                <td>10-05</td>
                                <td>15-05</td>
                                <td>20-05</td>
                                <td>20-05</td>
                                <td>1.000.000</td>
                                <td>1.505.000</td>
                                <td>1.000.000</td>
                            </tr>
                            <tr>
                                <td>SS #15</td>
                                <td>Just-Female</td>
                                <td>Supplier</td>
                                <td>10-05</td>
                                <td>15-05</td>
                                <td>20-05</td>
                                <td>20-05</td>
                                <td>1.000.000</td>
                                <td>1.505.000</td>
                                <td>1.000.000</td>
                            </tr>
                        </tbody>
                    </table>
                </section>
            </div>
            <!--TABLE END -->
        </div>
    </div>


</div>