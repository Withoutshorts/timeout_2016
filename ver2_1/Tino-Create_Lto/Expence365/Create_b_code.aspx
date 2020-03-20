<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Create_b_code.aspx.cs" Inherits="Create_b_code" EnableEventValidation="false" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head  runat="server">
    <meta charset="utf-8" />

    
    <link rel="stylesheet" type="text/css" href="Create_b.css" />


     <!-- iCONS -->
    <link rel="stylesheet" href="path/to/font-awesome/css/font-awesome.min.css" />
    <script src="https://kit.fontawesome.com/bbf7486ee3.js"></script>

    <!-- Plugin CSS -->
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/r/bs-3.3.5/dt-1.10.9/datatables.min.css" />

    <!-- New: Bootstrap Date-Picker Plugin -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/css/bootstrap-datepicker3.css" />

    <!--New: For jquery template-->
    <script src="http://code.jquery.com/jquery-latest.min.js"></script>
    <script src="http://ajax.microsoft.com/ajax/jquery.templates/beta1/jquery.tmpl.min.js"></script>

    <!--New: Bootstrap muliselect dropdown-->
    <link rel="stylesheet" href="http://netdna.bootstrapcdn.com/bootstrap/3.0.2/css/bootstrap.min.css" />
    <script src="http://netdna.bootstrapcdn.com/bootstrap/3.0.2/js/bootstrap.min.js"></script>

    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.13/css/bootstrap-multiselect.css" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-multiselect/0.9.13/js/bootstrap-multiselect.js"></script>


    <!-- New: Bootstrap Date-Picker Plugin -->
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-datepicker/1.4.1/js/bootstrap-datepicker.min.js"></script>

    <!-- jQuery validation -->
    <script type="text/javascript" src="https://cdn.jsdelivr.net/jquery.validation/1.15.1/jquery.validate.min.js"></script>

    <!-- BarDragger. draggable bar element in between panels -->
    <script src="https://static.codepen.io/assets/editor/pen/index-85a7604f435488c27ecf9076db52aa007293ef1665c535f985df3891dc5b1797.js"></script>
    <script src="https://static.codepen.io/assets/common/browser_support-f97f3ef9d187e4d0436e96e4d002603225517239e71f56f5a0441cce3a0d5be4.js"></script>
    <!-- 2005, 2014 jQuery Foundation  -->
    <script src="https://static.codepen.io/assets/common/everypage-eb67a83772e5fbc39195f098ad2e9fafa910022bcae0317bc5b3873698665a9e.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/vendor-1aaaa35c84754a5df90e.chunk.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/2-1aaaa35c84754a5df90e.chunk.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/everypage-1aaaa35c84754a5df90e.js"></script>
    <script src="https://static.codepen.io/assets/packs/js/processorRouter-1aaaa35c84754a5df90e.js"></script>

    <!--  jQuery UI - v1.11.3 - 2015-03-05 -->
    <script src="https://static.codepen.io/assets/editor/global/commonLibs-5aa21c051f554186721dfb2e22efa0262d769cfc8be7ee60a1c6e0e44c2bcd81.js"></script>
    <script src="https://static.codepen.io/assets/editor/global/codemirror-41b003ba7076ac1285f386080f36114ec6e4c72796832faf66404c01df1eb104.js"></script>
    <script src="https://static.codepen.io/assets/libs/emmet-codemirror-plugin-222bd3cb88c16bb29433e34064a6dce2845b15c040718116c240719eaafc143f.js"></script>
  



    <title>Expence365 - creation of new buyer</title>

</head>
<body>

    <section class="form-box" >
        
            <div class="container">
                
                <div class="row">
                    <div class="col-sm-10 col-sm-offset-1 col-md-8 col-md-offset-2 col-lg-6 col-lg-offset-3 form-wizard">
					
					<!-- Form Wizard -->
                    	<form role="form" runat="server" method="POST">
                            <button runat="server"  onserverclick="Redirect_til_home" class="homeBtn"><i class="fa fa-home"></i></button>
                    		<h2>Sign Up - Expence365</h2>
                    		<p>Udfyld alle felterne forneden før du går videre til næste side.</p>
							
						<!-- Form progress -->
                    		<div class="form-wizard-steps form-wizard-tolal-steps-4">
                    			<div class="form-wizard-progress">
                    			<div class="form-wizard-progress-line" data-now-value="12.25" data-number-of-steps="4" style="width: 12.25%;"></div>
                    			</div>
							<!-- Step 1 -->
                    			<div class="form-wizard-step active">
                    				<div class="form-wizard-step-icon"><i class="fa fa-building" aria-hidden="true"></i></div>
                    				<p>Firma Info</p>
                    			</div>
							<!-- Step 1 -->
								
							<!-- Step 2 -->
                    			<div class="form-wizard-step">
                    				<div class="form-wizard-step-icon"><i class="fa fa-location-arrow" aria-hidden="true"></i></div>
                    				<p>Address</p>
                    			</div>
							<!-- Step 2 -->
								
							<!-- Step 3 -->
								<div class="form-wizard-step">
                    				<div class="form-wizard-step-icon"><i class="fa fa-user" aria-hidden="true"></i></div>
                    				<p>Contact Info</p>
                    			</div>
							<!-- Step 3 -->
								
							<!-- Step 4 -->
								<div class="form-wizard-step">
                    				<div class="form-wizard-step-icon"><i class="fa fa-check-circle" aria-hidden="true"></i></div>
                    				<p>Completet</p>
                    			</div>
							<!-- Step 4 -->
                    		</div>
						<!-- Form progress -->
                    		
							
						<!-- Form Step 1 -->
                            <fieldset>
                                <h4>Company Information : <span>Step 1 - 4</span></h4>
								<div class="form-group">
                    			    <label>Company Name: <span>*</span></label>
                                    <input type="text" runat="server" id = "CompanyName" name="Company Name" placeholder="Company Name" class="form-control required"/>
                                </div>
                                <div class="form-group">
                    			    <label>Company initials: <span>*</span></label>
                                    <input type="text" runat="server" id = "CompanyInitials" name="Company initials" placeholder="Company initials  (max 5 letters.)" class="form-control required"  maxlength="5"/>
                                </div>
                                <div class="form-group">
                    			    <label>CVR Number: <span>*</span></label>
                                    <input type="number" runat="server" id = "CVRNumber" name="CVR Number" placeholder="CVR Number" class="form-control required"/>
                                </div>
								<div class="form-group">
                    			    <label>EAN Number:</label>
                                    <input type="number"  runat="server" id = "EANNumber" name="EAN Number" placeholder="EAN Number" class="form-control"/>
                                </div>               
                                <div class="form-group">
                    			    <label>Company Phone: <span>*</span></label>
                                    <input type="number" runat="server" id = "CompanyPhone" name="Phone" placeholder="Phone" class="form-control required"/>
                                </div>
								
                                <div class="form-wizard-buttons">
                                    <button type="button" class="btn btn-next">Next</button>
                                </div>
                            </fieldset>
                    	<!-- Form Step 1 -->

						<!-- Form Step 2 -->
                             <fieldset>
                                <h4>Company Address: <span>Step 2 - 4</span></h4>
                                <div class="form-group">
                    			    <label>Company Address: <span>*</span></label>
                                    <input type="text" runat="server" id = "Address" name="Address" requiredvalue="@" placeholder="Address" class="form-control required"/>
                                </div>
								<div class="form-group">
                    			    <label>Company Zip Code: <span>*</span></label>
                                    <input type="number" runat="server" id = "ZipCode" name="Zip Code" placeholder="Zip Code" class="form-control required"/>
                                </div>
								<div class="form-group">
                    			    <label>Company City: <span>*</span></label>
                                    <input type="text" runat="server" id = "City" name="City" placeholder="City" class="form-control required"/>
                                </div>
								
								<div class="form-group">
                                <label>Company Country: <span>*</span></label>
                                    <input type="text" runat="server" id = "Country" name="Country" placeholder="Country" class="form-control required"/>
                                </div>
					
                                <div class="form-wizard-buttons">
                                    <button type="button" class="btn btn-previous">Previous</button>
                                    <button type="button" class="btn btn-next">Next</button>
                                </div>
                            </fieldset>
						<!-- Form Step 2 -->

						<!-- Form Step 3 -->
                            <fieldset>
                                <h4>Contact Information: <span>Step 3 - 4</span></h4>
								<div class="form-group">
                    			    <label>Contact Person: <span>*</span></label>
                                    <input type="text" runat="server" id="ContactPerson" runat="server" name="Contact Person" placeholder="Contact Person" class="form-control required" />
                                </div>
                                <div class="form-group">
                    			    <label>Email: <span>*</span></label>
                                    <input type="text" runat="server" id = "Email" name="Email" placeholder="Email" class="form-control required"/>
                                </div>
								<div class="form-group">
                    			    <label>Phone: <span>*</span></label>
                                    <input type="number" runat="server" id = "ContactPersonPhone" name="Phone" placeholder="Phone" class="form-control required"/>
                                </div>
								
                                <div class="form-wizard-buttons">
                                    <button type="button" class="btn btn-previous">Previous</button>
                                    <button type="button" class="btn btn-next">Next</button>
                                </div>
                            </fieldset>
						<!-- Form Step 3 -->
							
						<!-- Form Step 4 -->
							<fieldset>
                                <h4>Completmet<span>Step 4 - 4</span></h4>
								<div style="clear:both;"></div>

                                <p>You have now completet your sign up </p>
                                <p>When u click submit all your information will be saved in our database and you will resive and email on " email " with a link so you can create your password.</p>
                                <p>If the email is incorrect please go back to previous site and correct it.</p>
                                
                                <p align="center">SUPPORT</p>
                                <p align="center">+45 25 36 55 00</p>
                                <p align="center">support@outzource.dk</p>
							
                                <div class="form-wizard-buttons">
                                    <button type="button" class="btn btn-previous">Previous</button>
                                    <asp:Button class="btn btn-submit" ID="btnSubmit" Text="Submit" runat="server" OnClick="BtnUpload_Click" />
                                 

                                </div>
                            </fieldset>
						<!-- Form Step 4 -->

                    	
                    	</form>     
					<!-- Form Wizard -->
                    </div>
                </div>
                    
            </div>
        </section>
    <script src="Create_b.js"></script>
</body>


<script>

   
    function scroll_to_class(element_class, removed_height) {
        var scroll_to = $(element_class).offset().top - removed_height;
        if ($(window).scrollTop() != scroll_to) {
            $('.form-wizard').stop().animate({ scrollTop: scroll_to }, 0);
        }
    }

    // Changes the numbers/icons for what side u are on.
    function bar_progress(progress_line_object, direction) {
        var number_of_steps = progress_line_object.data('number-of-steps');
        var now_value = progress_line_object.data('now-value');
        var new_value = 0;
        if (direction == 'right') {
            new_value = now_value + (100 / number_of_steps);
        }
        else if (direction == 'left') {
            new_value = now_value - (100 / number_of_steps);
        }
        progress_line_object.attr('style', 'width: ' + new_value + '%;').data('now-value', new_value);
    }

    function scroll_to_class(element_class, removed_height) {
        var scroll_to = $(element_class).offset().top - removed_height;
        if ($(window).scrollTop() != scroll_to) {
            $('.form-wizard').stop().animate({ scrollTop: scroll_to }, 0);
        }
    }

    function bar_progress(progress_line_object, direction) {
        var number_of_steps = progress_line_object.data('number-of-steps');
        var now_value = progress_line_object.data('now-value');
        var new_value = 0;
        if (direction == 'right') {
            new_value = now_value + (100 / number_of_steps);
        }
        else if (direction == 'left') {
            new_value = now_value - (100 / number_of_steps);
        }
        progress_line_object.attr('style', 'width: ' + new_value + '%;').data('now-value', new_value);
    }



    /*
        Form
    */
    //show first form
    $('.form-wizard fieldset:first').fadeIn('slow');

    $('.form-wizard .required').on('focus', function () {
        $(this).removeClass('input-error');
    });

    // next step
    $('.form-wizard .btn-next').on('click', function () {
        var parent_fieldset = $(this).parents('fieldset');
        var next_step = true;
        // navigation steps / progress steps
        var current_active_step = $(this).parents('.form-wizard').find('.form-wizard-step.active');
        var progress_line = $(this).parents('.form-wizard').find('.form-wizard-progress-line');

        // fields validation
        parent_fieldset.find('.required').each(function () {
            if ($(this).val() == "") {
                $(this).addClass('input-error');
                next_step = false;
            }
            else {
                $(this).removeClass('input-error');
            }
        });
        // fields validation


        if (next_step) {
            parent_fieldset.fadeOut(400, function () {
                // change icons
                current_active_step.removeClass('active').addClass('activated').next().addClass('active');
                // progress bar
                bar_progress(progress_line, 'right');
                // show next step
                $(this).next().fadeIn();
                // scroll window to beginning of the form
                scroll_to_class($('.form-wizard'), 20);
            });
        }

    });


    // previous step
    $('.form-wizard .btn-previous').on('click', function () {
        // navigation steps / progress steps
        var current_active_step = $(this).parents('.form-wizard').find('.form-wizard-step.active');
        var progress_line = $(this).parents('.form-wizard').find('.form-wizard-progress-line');

        $(this).parents('fieldset').fadeOut(400, function () {
            // change icons
            current_active_step.removeClass('active').prev().removeClass('activated').addClass('active');
            // progress bar
            bar_progress(progress_line, 'left');
            // show previous step
            $(this).prev().fadeIn();
            // scroll window to beginning of the form
            scroll_to_class($('.form-wizard'), 20);
        });
    });


  


    // image uploader scripts 

    var $dropzone = $('.image_picker'),
        $droptarget = $('.drop_target'),
        $dropinput = $('#inputFile'),
        $dropimg = $('.image_preview'),
        $remover = $('[data-action="remove_current_image"]');

    $dropzone.on('dragover', function () {
        $droptarget.addClass('dropping');
        return false;
    });

    $dropzone.on('dragend dragleave', function () {
        $droptarget.removeClass('dropping');
        return false;
    });

    $dropzone.on('drop', function (e) {
        $droptarget.removeClass('dropping');
        $droptarget.addClass('dropped');
        $remover.removeClass('disabled');
        e.preventDefault();

        var file = e.originalEvent.dataTransfer.files[0],
            reader = new FileReader();

        reader.onload = function (event) {
            $dropimg.css('background-image', 'url(' + event.target.result + ')');
        };

        console.log(file);
        reader.readAsDataURL(file);

        return false;
    });

    $dropinput.change(function (e) {
        $droptarget.addClass('dropped');
        $remover.removeClass('disabled');
        $('.image_title input').val('');

        var file = $dropinput.get(0).files[0],
            reader = new FileReader();

        reader.onload = function (event) {
            $dropimg.css('background-image', 'url(' + event.target.result + ')');
        }

        reader.readAsDataURL(file);
    });

    $remover.on('click', function () {
        $dropimg.css('background-image', '');
        $droptarget.removeClass('dropped');
        $remover.addClass('disabled');
        $('.image_title input').val('');
    });

    $('.image_title input').blur(function () {
        if ($(this).val() != '') {
            $droptarget.removeClass('dropped');
        }
    });

</script>



</html>

