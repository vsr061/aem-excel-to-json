$(document).ready(function () {

    $('#xls').change(function(){
           $('#selectedFile').text($('#xls').val().replace(/C:\\fakepath\\/i, ''));
    });

    $("#btnSubmit").click(function (event) {

        //validate input fields
        if ($("#xls").val().length > 1 && $("#destPath").val().length > 1) {

            var filename = $("#xls").val();

            // Use a regular expression to trim everything before final dot
            var extension = filename.replace(/^.*\./, '');
    
            // If there is no dot anywhere in filename, we would have extension == filename
            if (extension == filename) {
                extension = '';
            } else {
                // if there is an extension, we convert to lower case
                // (N.B. this conversion will not effect the value of the extension
                // on the file upload.)
                extension = extension.toLowerCase();
            }


            if(extension != 'xls' && extension != 'xlsx')
            {

                alert("Only xls or xlsx formats are allowed!");
                event.preventDefault();
                return;
            }
    

            
            //stop submit the form, we will post it manually.
            event.preventDefault();

            // Get form
            var form = $('#fileUploadForm')[0];

            // Create an FormData object 
            var data = new FormData(form);

            data.append("destPath", $("#destPath").val());

            // disabled the submit button
            $("#btnSubmit").prop("disabled", true);

            $(".loading").removeClass("loading--hide").addClass("loading--show");
            $(".result label").hide();

            $.ajax({
                type: "POST",
                enctype: 'multipart/form-data',
                url: "/apps/get/json/from/xls",
                data: data,
                processData: false,
                contentType: false,
                cache: false,
                success: function (data) {

                    $(".result label").text(data);
                    $(".result label").show();
                    $(".loading").removeClass("loading--show").addClass("loading--hide");
                    $("#btnSubmit").prop("disabled", false);

                },
                error: function (e) {

                    $(".result label").text(e.responseText);
                    $(".result label").show();
                    $(".loading").removeClass("loading--show").addClass("loading--hide");
                    $("#btnSubmit").prop("disabled", false);

                }
            });
        }
        else
        {
            alert("Please, fill the mandatory fields");
            // Cancel the form submission
            event.preventDefault();
            return;
        }



    });

});