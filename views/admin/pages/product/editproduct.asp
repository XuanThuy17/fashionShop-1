<!-- #include file="aspuploader/include_aspuploader.asp" -->

<% 'code here
Dim connDB
set connDB = Server.CreateObject("ADODB.Connection")
Dim strConnection
strConnection = "Provider=SQLOLEDB.1;Data Source=THUY092\SQLEXPRESS;Database=shop;User Id=sa;Password=123"
connDB.ConnectionString = strConnection
connDB.Open()

Set ID_product = Request.QueryString("id")
%> 

<!DOCTYPE html>
<html lang="en">

<head>
    <!-- <link rel="stylesheet" href="./css/modal.css"> -->
    <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <title>Star Admin2 </title>
    <!-- plugins:css -->
    <link rel="stylesheet" href="../../vendors/feather/feather.css">
    <link rel="stylesheet" href="../../vendors/mdi/css/materialdesignicons.min.css">
    <link rel="stylesheet" href="../../vendors/ti-icons/css/themify-icons.css">
    <link rel="stylesheet" href="../../vendors/typicons/typicons.css">
    <link rel="stylesheet" href="../../vendors/simple-line-icons/css/simple-line-icons.css">
    <link rel="stylesheet" href="../../vendors/css/vendor.bundle.base.css">
    <!-- endinject -->
    <!-- Plugin css for this page -->
    <!-- End plugin css for this page -->
    <!-- inject:css -->
    <link rel="stylesheet" href="../../css/vertical-layout-light/style.css">
    <!-- endinject -->
    <link rel="shortcut icon" href="../../images/favicon.png" />
</head>
<style>
    .swal2-confirm.swal2-styled {
      background-color: rgb(48, 133, 214) !important;
      color: #fff !important;
      font-weight: 400 !important;
    }
    .m-10-20 {
        margin: 10px 20px;
    }
</style>

<body>
    <div class="container-scroller">
      <!-- #include file="../../partials/_header.asp" --> 

      <div class="container-fluid page-body-wrapper">
        <!-- #include file="../../partials/_settings-panel.asp" -->

        <!-- #include file="../../partials/_sidebar.asp" -->

        <div class="main-panel">
          <div class="content-wrapper">
            <div class="row">
                <div class="col-12 grid-margin stretch-card">
              <div class="card">
                <div class="card-body">
                    <h4 class="card-title">Edit a product</h4>
                    <p class="card-description">
                        Enter information for products
                    </p>
                    <form class="forms-sample" method="POST">
                        <%
                        Set cmdPrep = Server.CreateObject("ADODB.Command")
                        cmdPrep.ActiveConnection = connDB
                        cmdPrep.CommandType = 1
                        cmdPrep.Prepared = True
                        cmdPrep.CommandText = "select product.*, sale_percent, start_day, end_day, brand from product join discount on product.ID_product = discount.ID_product join brand on product.ID_product = brand.ID_product where product.ID_product = "&ID_product

                        Set Result = cmdPrep.execute
                        do while not Result.EOF
                        %>
                        <div class="form-group">
                            <label for="exampleInputName1">Name product</label>
                            <input name="nameProduct" value="<%=Result("name")%>" type="text" class="form-control" id="exampleInputName1" placeholder="Name product" required>
                        </div>
                        <div class="form-group">
                            <label for="exampleInputName1">Brand</label>
                            <input name="brandProduct" value="<%=Result("brand")%>" type="text" class="form-control" id="exampleInputName1" placeholder="Brand product" required>
                        </div>
                        <div class="form-group">
                            <label for="exampleInputEmail3">Description</label>
                            <input name="description" value="<%=Result("description")%>" type="text" class="form-control" id="exampleInputEmail3" placeholder="Description" required>
                        </div>
                        <div class="form-group">
                            <label for="exampleSelectGender">Species</label>
                            <select name="species" class="form-control" id="selectSpecies" required>
                                <option><%=Result("species")%></option>
                                <option>Dresses</option>
                                <option>Bodysuits</option>
                                <option>Hoodies & Sweats</option>
                                <option>Jackets & Coats</option>
                                <option>Jeans</option>
                                <option>Pants & Leggings</option>
                                <option>Rompers & Jumpsuits</option>
                                <option>Shirts & Blouses</option>
                                <option>Shirts</option>
                                <option>Sweaters & Knits</option>
                                <option>Bras & Panties</option>
                            </select>
                        </div>
                        <div class="form-group">
                            <label for="exampleInputPassword4">Price</label>
                            <div class="input-group">
                                <div class="input-group-prepend">
                                    <span class="input-group-text bg-primary text-white">$</span>
                                </div>
                                <input name="price" value="<%=Result("price")%>" type="number" class="form-control" aria-label="Amount (to the nearest dollar)" required>
                                <div class="input-group-append">
                                    <span class="input-group-text">.00</span>
                                </div>
                            </div>
                        </div>
                        <div class="form-group">
                            <label for="">Color</label>
                            <div class="input-group">
                            <select name="color" id="add-color-input" class="form-control" id="exampleSelectGender" required>
                                <option value="1">BLACK</option>
                                <option value="2">WHITE</option>
                                <option value="3">GREY</option>
                                <option value="4">BLUE</option>
                                <option value="5">RED</option>
                                <option value="6">YELLOW</option>
                                <option value="7">ORANGE</option>
                            </select>
                            <div class="input-group-append">
                                <a id="add-color-btn" style="border-radius: 2px;" class="btn btn-sm btn-primary" type="button">Add</a>
                            </div>
                            </div>
                        </div>
                        
                        <div id="add-size">
                            <%
                            cmdPrep.CommandText = "select p.ID_product, color.color, color.ID_color from color join product_size_color p on p.ID_color = color.ID_color where p.ID_product = "&ID_product&" group by p.ID_product, color.color, color.ID_color"

                            Set ResultColor = cmdPrep.execute
                            do while not ResultColor.EOF
                            %>
                                <div class="form-group color-block" style="border: 1px solid #0003; padding: 5px;">
                                <label class="color-size-name" style="min-width: 10em;" for="exampleInputPassword4"><%=ResultColor("color")%></label>
                                <a class="del-color-btn" style="padding: 10px; float: right; cursor: pointer;"><i class="mdi mdi-delete-forever"></i></a>
                                <input class="input-id_color" style="display: none;" type="number" value="<%=ResultColor("ID_color")%>">
                                
                                <div class="col-md-6">
                                    <div class="form-group row">
                                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">S</label>
                                        <div class="col-sm-9">
                                        <%
                                        cmdPrep.CommandText = "select size.*, color.color, quantity from size join product_size_color p on p.ID_size = size.ID_size join color on color.ID_color = p.ID_color where p.ID_product = "&ID_product&" and size.ID_size = 1 and color.ID_color = "&ResultColor("ID_color")

                                        Set ResultSize = cmdPrep.execute
                                        if ResultSize.EOF then 
                                        %>
                                        <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityS" placeholder="Enter quantity" type="number" class="form-control">
                                        <%
                                        end if
                                        do while not ResultSize.EOF
                                        %>
                                            <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityS" placeholder="Enter quantity" value="<%=ResultSize("quantity")%>" type="number" class="form-control">
                                        <%
                                        ResultSize.MoveNext 
                                        loop
                                        %>
                                        </div>
                                    </div>
                                </div>

                                <div class="col-md-6">
                                    <div class="form-group row">
                                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">M</label>
                                        <div class="col-sm-9">
                                        <%
                                        cmdPrep.CommandText = "select size.*, color.color, quantity from size join product_size_color p on p.ID_size = size.ID_size join color on color.ID_color = p.ID_color where p.ID_product = "&ID_product&" and size.ID_size = 2 and color.ID_color = "&ResultColor("ID_color")

                                        Set ResultSize = cmdPrep.execute
                                        if ResultSize.EOF then 
                                        %>
                                        <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityM" placeholder="Enter quantity" type="number" class="form-control">
                                        <%
                                        end if
                                        do while not ResultSize.EOF
                                        %>
                                            <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityM" placeholder="Enter quantity" value="<%=ResultSize("quantity")%>" type="number" class="form-control">
                                        <%
                                        ResultSize.MoveNext 
                                        loop
                                        %>
                                        </div>
                                    </div>
                                </div>
                                
                                <div class="col-md-6">
                                    <div class="form-group row">
                                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">L</label>
                                        <div class="col-sm-9">
                                        <%
                                        cmdPrep.CommandText = "select size.*, color.color, quantity from size join product_size_color p on p.ID_size = size.ID_size join color on color.ID_color = p.ID_color where p.ID_product = "&ID_product&" and size.ID_size = 3 and color.ID_color = "&ResultColor("ID_color")

                                        Set ResultSize = cmdPrep.execute
                                        if ResultSize.EOF then 
                                        %>
                                        <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityL" placeholder="Enter quantity" type="number" class="form-control">
                                        <%
                                        end if
                                        do while not ResultSize.EOF
                                        %>
                                            <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityL" placeholder="Enter quantity" value="<%=ResultSize("quantity")%>" type="number" class="form-control">
                                        <%
                                        ResultSize.MoveNext 
                                        loop
                                        %>
                                        </div>
                                    </div>
                                </div>

                                <div class="col-md-6">
                                    <div class="form-group row">
                                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">XL</label>
                                        <div class="col-sm-9">
                                        <%
                                        cmdPrep.CommandText = "select size.*, color.color, quantity from size join product_size_color p on p.ID_size = size.ID_size join color on color.ID_color = p.ID_color where p.ID_product = "&ID_product&" and size.ID_size = 4 and color.ID_color = "&ResultColor("ID_color")

                                        Set ResultSize = cmdPrep.execute
                                        if ResultSize.EOF then 
                                        %>
                                        <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityXL" placeholder="Enter quantity" type="number" class="form-control">
                                        <%
                                        end if
                                        do while not ResultSize.EOF
                                        %>
                                            <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityXL" placeholder="Enter quantity" value="<%=ResultSize("quantity")%>" type="number" class="form-control">
                                        <%
                                        ResultSize.MoveNext 
                                        loop
                                        %>
                                        </div>
                                    </div>
                                </div>

                                <div class="col-md-6">
                                    <div class="form-group row">
                                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">XXL</label>
                                        <div class="col-sm-9">
                                        <%
                                        cmdPrep.CommandText = "select size.*, color.color, quantity from size join product_size_color p on p.ID_size = size.ID_size join color on color.ID_color = p.ID_color where p.ID_product = "&ID_product&" and size.ID_size = 5 and color.ID_color = "&ResultColor("ID_color")

                                        Set ResultSize = cmdPrep.execute
                                        if ResultSize.EOF then 
                                        %>
                                        <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityXXL" placeholder="Enter quantity" type="number" class="form-control">
                                        <%
                                        end if
                                        do while not ResultSize.EOF
                                        %>
                                            <input class="quantityColor<%=ResultColor("ID_color")%>" name="quantityXXL" placeholder="Enter quantity" value="<%=ResultSize("quantity")%>" type="number" class="form-control">
                                        <%
                                        ResultSize.MoveNext 
                                        loop
                                        %>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <%
                            ResultColor.MoveNext
                            loop
                            %>
                        </div>

                        <div class="form-group">
                        <label>File upload</label>
                        <input type="file" name="img[]" class="file-upload-default">
                        <div class="input-group col-xs-12">
                            <input type="text" class="form-control file-upload-info" disabled placeholder="Upload Image" required>
                            <span class="input-group-append">
                                <%
                                Dim uploader
                                Set uploader = new AspUploader
                                uploader.MaxSizeKB = 10240
                                uploader.Name = "myuploader"
                                uploader.MultipleFilesUpload = true
                                uploader.AllowedFileExtensions = "*.jpg,*.png"
                                uploader.SaveDirectory = "/fashionShop/resources/imgProduct"
                                uploader.InsertText = "Upload" 

                                ' response.write (uploader.GetString())   

                                %>
                                
                            </span>
                        </div>
                        </div>
                        
                        <ol id="filelist">
                        </ol>	

                        <div class="form-group">
                            <label for="exampleInputName1">Sale Percent</label>
                            <input name="salePercent" value="<%=Result("sale_percent")%>" type="number" class="form-control" id="exampleInputName1" placeholder="Enter sale percent" required>
                        </div>
                        <div class="row">
                        <div class="form-group col-md-6">
                            <label for="exampleInputName1">Start Day</label>
                            <input name="startDay" type="date" class="form-control" id="exampleInputName1" placeholder="Brand product" required>
                            <input name="startDayOld" type="text" value="<%=Result("start_day")%>" class="form-control" id="exampleInputName1" placeholder="Brand product" required>
                        </div>
                        <div class="form-group col-md-6">
                            <label for="exampleInputName1">End Day</label>
                            <input name="endDay" type="date" class="form-control" id="exampleInputName1" placeholder="Brand product" required>
                            <input name="endDayOld" type="text" value="<%=Result("end_day")%>" class="form-control" id="exampleInputName1" placeholder="Brand product" required>
                        </div>
                        </div>
                        
                        <a class="submitAdd btn btn-primary me-2">Add</a>
                    <%
                        Result.MoveNext
                        loop
                    %>
                    </form>
                </div>
              </div>
            </div>
            </div>
          </div>
        <!-- content-wrapper ends -->
        <!-- partial:partials/_footer.html -->
        <footer class="footer">
          <div class="d-sm-flex justify-content-center justify-content-sm-between">
            <span class="text-muted text-center text-sm-left d-block d-sm-inline-block">Premium <a href="https://www.bootstrapdash.com/" target="_blank">Bootstrap admin template</a> from BootstrapDash.</span>
            <span class="float-none float-sm-right d-block mt-1 mt-sm-0 text-center">Copyright © 2021. All rights reserved.</span>
          </div>
        </footer>
        <!-- partial -->
      </div>
      <!-- main-panel ends -->
    </div>
    <!-- page-body-wrapper ends -->
    </div>
    </div>
    <!-- container-scroller -->
    <!-- #include file="../../js/mainJs.asp" -->

    <!-- AJAX uploader -->
    <script type="text/javascript">
      var handlerurl='ajax-multiplefiles-handler.asp'
    </script>
    <script type="text/javascript">
    var listImage = []
    // function CuteWebUI_AjaxUploader_OnPostback()
    // {
    //     var uploader = document.getElementById("myuploader");
    //     var guidlist = uploader.value;
    
    //     //Send Request
    //     var xh;
    //     if (window.XMLHttpRequest)
    //         xh = new window.XMLHttpRequest();
    //     else
    //         xh = new ActiveXObject("Microsoft.XMLHTTP");
        
    //     xh.open("POST", handlerurl, false, null, null);
    //     xh.setRequestHeader("Content-Type", "application/x-www-form-urlencoded; charset=utf-8");
    //     xh.send("guidlist=" + guidlist);
    
    //     //call uploader to clear the client state
    //     uploader.reset();
    
    //     if (xh.status != 200)
    //     {
    //         alert("http error " + xh.status);
    //         setTimeout(function() { document.write(xh.responseText); }, 10);
    //         return;
    //     } 

    //     var filelist = document.getElementById("filelist");
    
    //     var list = eval(xh.responseText); //get JSON objects
    //     //Process Result:
    //     for (var i = 0; i < list.length; i++)
    //     {
    //         console.log(list[i].FileName)
    //         var item = list[i];
    //         var msg = "Processed: " + list[i].FileName;
    //         var li = document.createElement("li");
    //         li.innerHTML = msg;
    //         filelist.appendChild(li);
    //         listImage.push(list[i].FileName);
    //     }
    //     }
    //     var objListImage = []
        
    //     const createListImage = () => {
    //     if (listImage.length >= 4) {g
    //         objListImage.push({
    //         id_product: <%=ID_product%>,
    //         link1: listImage[0],
    //         link2: listImage[1],
    //         link3: listImage[2],
    //         link4: listImage[3]
    //         })
    //     } else if (listImage.length == 3){
    //         objListImage.push({
    //         id_product: <%=ID_product%>,
    //         link1: listImage[0],
    //         link2: listImage[1],
    //         link3: listImage[2],
    //         link4: ''
    //         })
    //     } else if (listImage.length == 2) {
    //         objListImage.push({
    //         id_product: <%=ID_product%>,
    //         link1: listImage[0],
    //         link2: listImage[1],
    //         link3: '',
    //         link4: ''
    //         })
    //     }
    //     var xmlhttp = new XMLHttpRequest();
    //     xmlhttp.open("GET", "/fashionShop/controllers/admin/editImageProduct.asp?id=" + objListImage[0].id_product + "&link1=" + objListImage[0].link1 + "&link2="+objListImage[0].link2+"&link3="+ objListImage[0].link3 + "&link4=" + objListImage[0].link4, true);
    //     // console.log(ID_product)
    //     xmlhttp.send();
    // }

    </script>
    <!-- End AJAX uploader -->

    <!-- Create block select color and quantity -->
    <script>
        var addColorBtn = document.getElementById('add-color-btn');
        var addColorInput = document.getElementById('add-color-input');
        var addSize = document.getElementById('add-size');
        var colorSizeName = document.querySelectorAll('.color-size-name');

        var delColorBtns
        var colorBlock
        var idColor
        var arrayProduct = []
        var colorExists = []
        
        // Click thêm color
        addColorBtn.addEventListener('click', () => {
            const selectedOption = addColorInput.options[addColorInput.selectedIndex];
            if (!colorExists.includes(selectedOption.text)) {
                colorExists.push(selectedOption.text);
                console.log(true)

                const str = selectedOption.text
                const strUpperCase = str.charAt(0).toUpperCase() + str.slice(1);
                addSize.innerHTML += `
                <div class="form-group color-block" style="border: 1px solid #0003; padding: 5px;">
                    <label class="color-size-name" style="min-width: 10em;" for="exampleInputPassword4">${strUpperCase}</label>
                    <a class="del-color-btn" style="padding: 10px; float: right; cursor: pointer;"><i class="mdi mdi-delete-forever"></i></a>
                    <input class="input-id_color" style="display: none;" type="number" value="${selectedOption.value}">
                    <div class="col-md-6">
                    <div class="form-group row">
                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">S</label>
                        <div class="col-sm-9">
                            <input class="quantityColor${selectedOption.value}" name="quantityS" placeholder="Enter quantity"  type="number" class="form-control">
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="form-group row">
                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">M</label>
                        <div class="col-sm-9">
                            <input class="quantityColor${selectedOption.value}" name="quantityM" placeholder="Enter quantity"  type="number" class="form-control">
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="form-group row">
                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">L</label>
                        <div class="col-sm-9">
                            <input class="quantityColor${selectedOption.value}" name="quantityL" placeholder="Enter quantity"  type="number" class="form-control">
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="form-group row">
                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">XL</label>
                        <div class="col-sm-9">
                            <input class="quantityColor${selectedOption.value}" name="quantityXL" placeholder="Enter quantity"  type="number" class="form-control">
                        </div>
                    </div>
                </div>
                <div class="col-md-6">
                    <div class="form-group row">
                        <label class="col-sm-3 col-form-label" style="padding-top: 0;">XXL</label>
                        <div class="col-sm-9">
                            <input class="quantityColor${selectedOption.value}" name="quantityXXL" placeholder="Enter quantity"  type="number" class="form-control">
                        </div>
                    </div>
                    </div>
                </div>
                `;
                // delColorBtns = document.querySelectorAll('.del-color-btn');
                // colorBlock = document.querySelectorAll('.color-block');
                // idColor = document.querySelectorAll('.input-id_color');
                // delColorBtns.forEach((btn, index) => {
                //     btn.addEventListener('click', () => {
                //         colorBlock[index].remove();
                //     })
                // })
                delFunc()
            } else {
                alert("Màu đã tồn tại")
            }
            
            
        })
        
        const delFunc = () => {
            delColorBtns = document.querySelectorAll('.del-color-btn');
            colorBlock = document.querySelectorAll('.color-block');
            idColor = document.querySelectorAll('.input-id_color');
            delColorBtns.forEach((btn, index) => {
                btn.addEventListener('click', () => {
                    colorBlock[index].remove();
                    delFunc()
                })
            })
        }
        delFunc()
        // const uploaderButton = document.querySelector('#myuploaderButton');
        // uploaderButton.classList.add('file-upload-browse')
        // uploaderButton.classList.add('btn')
        // uploaderButton.classList.add('btn-primary')

        const checkQuantity = (quantity, element, id_size) => {
            if (quantity > 0 && typeof(parseInt(quantity)) == 'number') {
                arrayProduct.push({
                id_product: <%=ID_product%>,
                id_color: parseInt(element.value),
                id_size: id_size,
                quantity: parseInt(quantity)
                })
            }
        }

        var nameProduct = ''
        var brandProduct = ''
        var description = ''
        var selectSprecies = ''
        var price = ''
        const alertEmptyForm = (text) => {
          if (text == '') {
            alert('Please enter enough information!');
          }
        }
        // check rỗng và add vào bảng product
        const checkEmptyForm = () => {
          nameProduct = document.querySelector('input[name="nameProduct"]').value;
          brandProduct = document.querySelector('input[name="brandProduct"]').value;
          description = document.querySelector('input[name="description"]').value;
          const species = document.querySelector('#selectSpecies');
          selectSprecies = species.options[species.selectedIndex].value;
          price = document.querySelector('input[name="price"]').value;
          alertEmptyForm(nameProduct)
          alertEmptyForm(brandProduct)
          alertEmptyForm(description)
          alertEmptyForm(price)


          nameProduct = nameProduct.replace(/ /g, "_");
          description = description.replace(/ /g, "_");
          brandProduct = brandProduct.replace(/ /g, "_");
          selectSprecies = selectSprecies.replace(/ /g, "_");
          selectSprecies = selectSprecies.replace(/&/g, "<>");
          price = price.replace(/ /g, "_");

          var xmlhttp = new XMLHttpRequest();
          xmlhttp.open("GET", "/fashionShop/controllers/admin/editProduct.asp?id=" + <%=ID_product%> + "&name=" + nameProduct + "&brand="+brandProduct+"&desc="+description + "&species="+selectSprecies+"&price=" + price, true);
          // console.log(ID_product)
          xmlhttp.send();
        }
        // add size & color & quantity
        const addSizeColorQuantity = () => {
            var callSuccess
            var xmlhttp = new XMLHttpRequest();
            // xmlhttp.onreadystatechange = function() {
            //     if (xmlhttp.status == 200) {
            //         callSuccess = true
            //     } else {
            //         callSuccess = false
            //         alert("False")
            //     }
            // };
            xmlhttp.open("GET", "/fashionShop/controllers/admin/delQuantity.asp?id_product=" + <%=ID_product%> + "&id_size=" + product.id_size + "&id_color=" + product.id_color, true);
            xmlhttp.send();

            arrayProduct.forEach((product) => {
                var xmlr = new XMLHttpRequest();
                xmlr.open("GET", "/fashionShop/controllers/admin/editQuantity.asp?id_product=" + <%=ID_product%> + "&id_size=" + product.id_size + "&id_color=" + product.id_color + "&quantity=" + product.quantity, true);
                
                xmlr.send();
            })

            // if (callSuccess) {
            //     arrayProduct.forEach((product) => {
            //         var xmlhttp = new XMLHttpRequest();
            //         xmlhttp.open("GET", "/fashionShop/controllers/admin/editQuantity.asp?id_product=" + <%=ID_product%> + "&id_size=" + product.id_size + "&id_color=" + product.id_color + "&quantity=" + product.quantity, true);
                    
            //         xmlhttp.send();
            //     })
            // } else {
            //     alert("False")
            // }
            // arrayProduct = []
        }
        // add salePercent
        const addSalePercent = () => {
            const postSalePercent = (id, start, end, sale) => {
                var xmlhttp = new XMLHttpRequest();
                xmlhttp.open("GET", "/fashionShop/controllers/admin/editSalePercent.asp?id_product=" + id + "&start=" + start + "&end=" + end + "&sale=" + sale, true);
                // console.log(ID_product)
                xmlhttp.send();
            }
            const salePercent = document.querySelector('input[name="salePercent"]').value;
            const startDay = document.querySelector('input[name="startDay"]').value;
            const startDayOld = document.querySelector('input[name="startDayOld"]').value;
            const endDay = document.querySelector('input[name="endDay"]').value;
            const endDayOld = document.querySelector('input[name="endDayOld"]').value;

            if (salePercent == '') {
                const salePercentValue = 0;
                postSalePercent(<%=ID_product%>, startDay, endDay, salePercentValue);
            } else {
                salePercentValue = salePercent;
                if (startDay == '' & endDay == '') {
                    postSalePercent(<%=ID_product%>, startDayOld, endDayOld, salePercentValue);
                } else {
                    postSalePercent(<%=ID_product%>, startDay, endDay, salePercentValue);
                }
            }
        }
        // click add số lượng vào mảng số lượng
        const submitBtn = document.querySelector('.submitAdd');
        submitBtn.addEventListener('click', () => {
            checkEmptyForm()
            idColor.forEach((e, index) => {
                const quantityS = document.querySelector(`.quantityColor${e.value}[name="quantityS"]`).value;
                id_size = 1
                checkQuantity(quantityS, e, id_size)
                const quantityM = document.querySelector(`.quantityColor${e.value}[name="quantityM"]`).value;
                id_size = 2
                checkQuantity(quantityM, e, id_size)
                const quantityL = document.querySelector(`.quantityColor${e.value}[name="quantityL"]`).value;
                id_size = 3
                checkQuantity(quantityL, e, id_size)
                const quantityXL = document.querySelector(`.quantityColor${e.value}[name="quantityXL"]`).value;
                id_size = 4
                checkQuantity(quantityXL, e, id_size)
                const quantityXXL = document.querySelector(`.quantityColor${e.value}[name="quantityXXL"]`).value;
                id_size = 5
                checkQuantity(quantityXXL, e, id_size)
            })
            addSizeColorQuantity()
            //   createListImage()
            addSalePercent()
            arrayProduct = []
        })
    </script> 

    <!-- End Create block select color and quantity -->

</body>

</html>