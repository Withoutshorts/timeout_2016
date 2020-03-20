


<!-- Aktivitet og job picker -->



<style>

        table, tr, td, .tablecolor 
        {
            color:black;
            padding:0 15px 10px 0px;
        }
    
    
    
        /* The Modal (background) */
        .modal {
            display: none; /* Hidden by default */
            position: fixed; /* Stay in place */
            z-index: 1; /* Sit on top */
            padding-top: 4px; /* Location of the box */
            left: 0;
            top: 0;
            width: 100%; /* Full width */
            height: 100%; /* Full height */
            overflow: auto; /* Enable scroll if needed */
            background-color: rgb(0,0,0); /* Fallback color */
            background-color: rgba(0,0,0,0.4); /* Black w/ opacity */
        }

        /* Modal Content */
        .modal-content {
            background-color: #fefefe;
            margin: auto;
            padding: 20px;
            border: 1px solid #888;
            width: 95%;
            height: 99%;
        }
        
        .picmodal:hover,
        .picmodal:focus {
        text-decoration: none;
        cursor: pointer;
        }   
        
        .hover:hover,
        .hover:focus {
        text-decoration: none;
        cursor: pointer;
        }

        .btn-paging {
            font-size:150%;
            margin-right: 25px;
            width:125px;
        }

        .btn-paging-active {
            font-size:150%;
            margin-right: 25px;
            width:125px;
            background-color:#dddddd;
        }
   
</style>






<div id="myModal_picker" class="modal" style="font-size:80%;">

    <!-- Modal content -->
    <div class="modal-content" style="overflow-y:scroll; overflow-x:hidden;">

        <link rel="stylesheet" href="//cdn.datatables.net/1.10.19/css/jquery.dataTables.min.css" />
        <script src="//cdn.datatables.net/1.10.19/js/jquery.dataTables.min.js"></script>

        <div class="row">
            <div class="col-lg-12">
                <table>
                    <tbody>
                        <tr>
                            <td>
                                <span style="font-size:150%; background-color:#bcbcbc; margin-left:25px; display:none;" id="jobsort_navn" class="btn btn-default hover"><b>Navn</b> <span class="fa fa-sort"></span></span>
                                <span style="font-size:150%; background-color:#bcbcbc; margin-left:25px; display:none;" id="aktsort_navn" class="btn btn-default hover"><b>Navn</b> <span class="fa fa-sort"></span></span>
                            </td>
                            <td>
                                <span style="font-size:150%; background-color:#bcbcbc; margin-left:25px;" class="btn btn-default backbtn hover"><b>Tilbage</b></span>
                            </td>   
                            <td style="text-align:center; width:100%; padding-right:240px;"><span style="font-size:150%;" id="overskriftspan"><b id="modal_overskrift"></b> <br /> <b style="font-size:9px;" id="modal_underskrift">&nbsp; Velg aktivitet &nbsp;</b> </span></td>

                            <td><span id="closebtn" style="font-size:150%; background-color:#db2b2b; margin-left:25px;" class="btn btn-default hover"><b style="color:white;">Lukk</b></span></td>

                        </tr>
                        </tbody>
                </table>
            </div>

        </div>

        <div class="row">
            <div id="jobSection" class="col-lg-12" style="height:750px;">
                <table id="job_table">

                    <thead style="display:none;">
                        <tr>
                            <th>1</th>
                            <th>2</th>
                            <th>3</th>
                            <th>4</th>
                            <th>5</th>
                        </tr>
                    </thead>
                                                
                    <tbody id="modal_job_SEL">

                    </tbody>
                </table>

                <table id="akt_table">

                    <thead style="display:none;">
                        <tr>
                            <th>1</th>
                            <th>2</th>
                            <th>3</th>
                            <th>4</th>
                        </tr>
                    </thead>
                                                
                    <tbody id="modal_akt_SEL">

                    </tbody>
                </table>
            </div>
        </div>
    </div>

</div>