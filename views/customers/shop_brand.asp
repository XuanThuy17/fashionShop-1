<!-- #include file="connect.asp" -->
<%
' ham lam tron so nguyen
    function Ceil(Number)
        Ceil = Int(Number)
        if Ceil<>Number Then
            Ceil = Ceil + 1
        end if
    end function

    function checkPage(cond, ret) 
        if cond=true then
            Response.write ret
        else
            Response.write ""
        end if
    end function
' trang hien tai
    page = Request.QueryString("page")
    ID_brand = Request.QueryString("brand")

    limit = 9

    if (trim(page) = "") or (isnull(page)) then
        page = 1
    end if

    offset = (Clng(page) * Clng(limit)) - Clng(limit)

    strSQL = "SELECT COUNT(product.ID_product) AS count FROM product inner join brand on product.ID_product = brand.ID_product where brand.brand = (select brand.brand from brand where brand.ID_brand = "&ID_brand&")"

    connDB.Open()
    Set CountResult = connDB.execute(strSQL)

    totalRows = CLng(CountResult("count"))

    Set CountResult = Nothing
' lay ve tong so trang
    pages = Ceil(totalRows/limit)
    ID_brand = Request.QueryString("brand")
%>
<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="description" content="">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Title  -->
    <title>Essence - Fashion Ecommerce Template</title>

    <link rel="icon" href="img/core-img/favicon.ico">
    <link rel="stylesheet" href="./css/add.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" integrity="sha512-iecdLmaskl7CVkqkXNQ/ZH/XLlvWZOJyj7Yy7tcenmpD1ypASozpmT/E0iPtmFIB46ZmdtAc9eNBvH0H/ZpiBw==" crossorigin="anonymous" referrerpolicy="no-referrer" />
    <!-- Core Style CSS -->
    <link rel="stylesheet" href="css/core-style.css">
    <link rel="stylesheet" href="style.css">
    
    <style>
        .nice-select {
            display: block !importan;
        }
    </style>
</head>

<body>
    <!-- ##### Header Area Start ##### -->
   
    <!-- #include file="header.asp" -->

    <!-- ##### Header Area End ##### -->

    <!-- #include file="cart.asp" -->

    <!-- ##### Breadcumb Area Start ##### -->
    <div class="breadcumb_area bg-img" style="background-image: url(img/bg-img/breadcumb.jpg);">
        <div class="container h-100">
            <div class="row h-100 align-items-center">
                <div class="col-12">
                    <div class="page-title text-center">
                        <h2>Shop</h2>
                    </div>
                </div>
            </div>
        </div>
    </div>
    <!-- ##### Breadcumb Area End ##### -->

    <!-- ##### Shop Grid Area Start ##### -->
    <section class="shop_grid_area section-padding-80">
        <div class="container">
            <div class="row">
                <div class="col-12 col-md-4 col-lg-3">
                    <div class="shop_sidebar_area">

                        <!-- ##### Single Widget ##### -->
                        <div class="widget catagory mb-50">
                            <!-- Widget Title -->
                            <h6 class="widget-title mb-30">Catagories</h6>

                            <!--  Catagories  -->
                            <div class="catagories-menu">
                                <ul id="menu-content2" class="menu-content collapse show">
                                    <!-- Single Item -->
                                    <li data-toggle="collapse" data-target="#clothing">
                                        <a href="#">clothing</a>
                                        <ul class="sub-menu collapse show" id="clothing">
                                            <li><a href="#">All</a></li>
                                            <li><a href="#">Bodysuits</a></li>
                                            <li><a href="#">Dresses</a></li>
                                            <li><a href="#">Hoodies &amp; Sweats</a></li>
                                            <li><a href="#">Jackets &amp; Coats</a></li>
                                            <li><a href="#">Jeans</a></li>
                                            <li><a href="#">Pants &amp; Leggings</a></li>
                                            <li><a href="#">Rompers &amp; Jumpsuits</a></li>
                                            <li><a href="#">Shirts &amp; Blouses</a></li>
                                            <li><a href="#">Shirts</a></li>
                                            <li><a href="#">Sweaters &amp; Knits</a></li>
                                        </ul>
                                    </li>
                                    <!-- Single Item -->
                                    <li data-toggle="collapse" data-target="#shoes" class="collapsed">
                                        <a href="#">shoes</a>
                                        <ul class="sub-menu collapse" id="shoes">
                                            <li><a href="#">All</a></li>
                                            <li><a href="#">Bodysuits</a></li>
                                            <li><a href="#">Dresses</a></li>
                                            <li><a href="#">Hoodies &amp; Sweats</a></li>
                                            <li><a href="#">Jackets &amp; Coats</a></li>
                                            <li><a href="#">Jeans</a></li>
                                            <li><a href="#">Pants &amp; Leggings</a></li>
                                            <li><a href="#">Rompers &amp; Jumpsuits</a></li>
                                            <li><a href="#">Shirts &amp; Blouses</a></li>
                                            <li><a href="#">Shirts</a></li>
                                            <li><a href="#">Sweaters &amp; Knits</a></li>
                                        </ul>
                                    </li>
                                    <!-- Single Item -->
                                    <li data-toggle="collapse" data-target="#accessories" class="collapsed">
                                        <a href="#">accessories</a>
                                        <ul class="sub-menu collapse" id="accessories">
                                            <li><a href="#">All</a></li>
                                            <li><a href="#">Bodysuits</a></li>
                                            <li><a href="#">Dresses</a></li>
                                            <li><a href="#">Hoodies &amp; Sweats</a></li>
                                            <li><a href="#">Jackets &amp; Coats</a></li>
                                            <li><a href="#">Jeans</a></li>
                                            <li><a href="#">Pants &amp; Leggings</a></li>
                                            <li><a href="#">Rompers &amp; Jumpsuits</a></li>
                                            <li><a href="#">Shirts &amp; Blouses</a></li>
                                            <li><a href="#">Shirts</a></li>
                                            <li><a href="#">Sweaters &amp; Knits</a></li>
                                        </ul>
                                    </li>
                                </ul>
                            </div>
                        </div>

                        <!-- ##### Single Widget ##### -->
                        <div class="widget price mb-50">
                            <!-- Widget Title -->
                            <h6 class="widget-title mb-30">Filter by</h6>
                            <!-- Widget Title 2 -->
                            <p class="widget-title2 mb-30">Price</p>

                            <div class="widget-desc">
                                <div class="slider-range">
                                    <div data-min="49" data-max="360" data-unit="$" class="slider-range-price ui-slider ui-slider-horizontal ui-widget ui-widget-content ui-corner-all" data-value-min="49" data-value-max="360" data-label-result="Range:">
                                        <div class="ui-slider-range ui-widget-header ui-corner-all"></div>
                                        <span class="ui-slider-handle ui-state-default ui-corner-all" tabindex="0"></span>
                                        <span class="ui-slider-handle ui-state-default ui-corner-all" tabindex="0"></span>
                                    </div>
                                    <div class="range-price">Range: $49.00 - $360.00</div>
                                </div>
                            </div>
                        </div>

                        <!-- ##### Single Widget ##### -->
                        <div class="widget color mb-50">
                            <!-- Widget Title 2 -->
                            <p class="widget-title2 mb-30">Color</p>
                            <div class="widget-desc">
                                <ul class="d-flex">
                                    <li><a href="#" class="color1"></a></li>
                                    <li><a href="#" class="color2"></a></li>
                                    <li><a href="#" class="color3"></a></li>
                                    <li><a href="#" class="color4"></a></li>
                                    <li><a href="#" class="color5"></a></li>
                                    <li><a href="#" class="color6"></a></li>
                                    <li><a href="#" class="color7"></a></li>
                                    <li><a href="#" class="color8"></a></li>
                                    <li><a href="#" class="color9"></a></li>
                                    <li><a href="#" class="color10"></a></li>
                                </ul>
                            </div>
                        </div>

                        <!-- ##### Single Widget ##### -->
                        <div class="widget brands mb-50">
                            <!-- Widget Title 2 -->
                            <p class="widget-title2 mb-30">Brands</p>
                            <div class="widget-desc">
                                <ul>
                                <%
                                Set Conn = Server.CreateObject("ADODB.Connection")
                                Conn.Open "Provider=SQLOLEDB.1;Data Source=THUY092\SQLEXPRESS;Database=shop;User Id=sa;Password=123"

                                sql = "SELECT MAX(ID_brand) AS ID_brand, brand FROM brand group by brand"

                                Set result_brand = Conn.execute(sql)
                                do while not result_brand.EOF
                                %>
                                    <li><a href="shop_brand.asp?brand=<%=result_brand("ID_brand")%>"><%=result_brand("brand")%></a></li>
                                <%
                                result_brand.MoveNext
                                loop
                                %>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <%
                Set cmdPrep = Server.CreateObject("ADODB.Command")
                cmdPrep.ActiveConnection = connDB
                cmdPrep.CommandType = 1
                cmdPrep.Prepared = True
                ' cmdPrep.CommandText = "SELECT * FROM PRODUCT INNER JOIN IMAGEPRODUCT ON PRODUCT.ID_PRODUCT = IMAGEPRODUCT.ID_PRODUCT ORDER BY ID_product OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
                if (NOT IsEmpty(Session("ID_user"))) then
                    cmdPrep.CommandText = "SELECT product.name, product.ID_product, new, sale_percent, end_day, brand, price, link1, link2 FROM product inner join discount on discount.ID_product = product.ID_product inner join brand on product.ID_product = brand.ID_product inner join imageProduct on product.ID_product = imageProduct.ID_product inner join users on users.ID_user = "&Session("ID_user")&" where brand.brand = (select brand from brand where brand.ID_brand = "&ID_brand&") GROUP BY product.name, product.ID_product, new, sale_percent, end_day, brand, price, link1, link2 ORDER BY ID_product OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
                else 
                    cmdPrep.CommandText = "SELECT product.name, product.ID_product, new, sale_percent, end_day, brand, price, link1, link2 FROM product inner join discount on discount.ID_product = product.ID_product inner join brand on product.ID_product = brand.ID_product inner join imageProduct on product.ID_product = imageProduct.ID_product where brand.brand = (select brand from brand where brand.ID_brand = "&ID_brand&") GROUP BY product.name, product.ID_product, new, sale_percent, end_day, brand, price, link1, link2 ORDER BY ID_product OFFSET ? ROWS FETCH NEXT ? ROWS ONLY"
                end if
                cmdPrep.parameters.Append cmdPrep.createParameter("offset",3,1, ,offset)
                cmdPrep.parameters.Append cmdPrep.createParameter("limit",3,1, , limit)


                Set Result = cmdPrep.execute
                cmdPrep.CommandText = "select product.ID_product, link1, link2 from product inner join imageProduct on product.ID_product = imageProduct.ID_product"
                Set ResultImg = cmdPrep.execute
                %>
                <div class="col-12 col-md-8 col-lg-9">
                    <div class="shop_grid_product_area">
                        <div class="row">
                            <div class="col-12">
                                <div class="product-topbar d-flex align-items-center justify-content-between">
                                    <!-- Total Products -->
                                    <div class="total-products">
                                        <p><span><%=totalRows%></span> products found</p>
                                    </div>
                                    <!-- Sorting -->
                                    <div class="product-sorting d-flex">
                                        <p>Sort by:</p>
                                        <form action="#" method="get">
                                            <select name="select" id="sortByselect">
                                                <option value="value">Highest Rated</option>
                                                <option value="value">Newest</option>
                                                <option value="value">Price: $$ - $</option>
                                                <option value="value">Price: $ - $$</option>
                                            </select>
                                            <input type="submit" class="d-none" value="">
                                        </form>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div class="row">
                        <% 
                            do while not Result.EOF
                        %>
                            <!-- Single Product -->
                            <div class="col-12 col-sm-6 col-lg-4">
                                <div class="single-product-wrapper">
                                    <!-- Product Image -->
                                    <div class="product-img">
                                        <img src="/fashionShop/resources/imgProduct/<%=Result("link1")%>" alt="">
                                        <input class="id_product" style="display: none;" value="<%=Result("ID_product")%>" >
                                        <!-- Hover Thumb -->
                                        <img class="hover-img" src="/fashionShop/resources/imgProduct/<%=Result("link2")%>" alt="">

                                        <!-- Product Badge -->
                                        <%
                                        percent = CInt(Result("sale_percent"))

                                        Dim currentDate
                                        currentDate = Date()

                                        Dim datee
                                        datee = FormatDateTime(Result("end_day"),2)

                                        if (CStr(datee) < CStr(currentDate)) then
                                            percent = 0
                                        end if
                                        if (percent > 0) then
                                        %>
                                        <div class="product-badge offer-badge">
                                            <span>-<%=percent%>%</span>
                                        </div>
                                        <% end if %>

                                        <% if (Result("new")) then%>
                                            <% if (percent > 0) then %>
                                                <div class="product-badge new-badge" style="margin-top: 3em;">
                                                    <span>New</span>
                                                </div>
                                            <% else %>
                                                <div class="product-badge new-badge">
                                                    <span>New</span>
                                                </div>
                                            <% end if %>
                                        <% end if %>

                                        <!-- Favourite -->
                                        <% if (NOT IsEmpty(Session("ID_user"))) then %>
                                        <div class="product-favourite">
                                        <%
                                            Set Conn = Server.CreateObject("ADODB.Connection")
                                            Conn.Open "Provider=SQLOLEDB.1;Data Source=THUY092\SQLEXPRESS;Database=shop;User Id=sa;Password=123"
                                            Dim sql 
                                            sql = "select * from favorite where ID_user = "&Session("ID_user")&" and ID_product = "&Result("ID_product")
                                            set rs = Conn.Execute(sql)

                                            if not rs.EOF then %>
                                                <a id="favo_<%=Result("ID_product")%>" href="#" class="favorite_btn favme fa fa-heart active "></a>
                                        <%  else   %>
                                                <a id="favo_<%=Result("ID_product")%>" href="#" class="favorite_btn favme fa fa-heart "></a>
                                        <%
                                            end if
                                        %>
                                        </div>
                                        <% else %>
                                        <div class="product-favourite">
                                            <a href="#" class="favorite_btn favme fa fa-heart"></a>
                                        </div>
                                        <% end if %>
                                        
                                    </div>

                                    <!-- Product Description -->
                                    <div class="product-description">
                                        <span><%=Result("brand")%></span>
                                        <a href="product_ex.asp?product=<%=CInt(Result("ID_product"))%>">
                                            <h6><%=Result("name")%></h6>
                                        </a>

                                        <%
                                        dim priceSale
                                        if (percent > 0) then 
                                            priceSale = CInt(Result("price")) - CInt(Result("price")) * percent / 100
                                        %>

                                        <p class="product-price"><span class="old-price">$<%=Result("price")%>.00</span> $<%=Ceil(priceSale)%>.00</p>
                                        
                                        <% else %>
                                        <p class="product-price">$<%=Result("price")%>.00</p>
                                        <% end if%>
                                        <!-- Hover Content -->
                                        <div class="hover-content">
                                            <!-- Add to Cart -->
                                            <div class="add-to-cart-btn">
                                                <a href="product_ex.asp?product=<%=CInt(Result("ID_product"))%>" class="btn essence-btn">Shop now</a>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        <%
                            Result.MoveNext
                            loop
                        %>
            
                        </div>
                    </div>
                    <!-- Pagination -->
                    <nav aria-label="navigation">
                        <ul class="pagination mt-50 mb-70">
                            <li class="page-item"><a class="page-link" href="#"><i class="fa fa-angle-left"></i></a></li>
                            <% 
                                if (pages > 1) then 
                                for i = 1 to pages

                            %>
                                <li class="page-item <%=checkPage(Clng(i)=Clng(page),"active-page")%>"><a class="page-link" href="shop.asp?page=<%=i%>"><%=i%></a></li>
                            <%
                                next
                                end if
                            %>
                            <li class="page-item"><a class="page-link" href="#"><i class="fa fa-angle-right"></i></a></li>
                        </ul>
                    </nav>
                </div>
            </div>
        </div>
    </section>
    <!-- ##### Shop Grid Area End ##### -->

    <!-- #include file="footer.asp" -->

    <!-- jQuery (Necessary for All JavaScript Plugins) -->
    <script src="js/jquery/jquery-2.2.4.min.js"></script>
    <script>
        var favoriteBtns = document.querySelectorAll(".favorite_btn");
        
        favoriteBtns.forEach((favoriteBtn, index) => {
            favoriteBtn.addEventListener('click', function() {
                var stringID = favoriteBtn.id    
                var id = stringID.charAt(stringID.length - 1)
                console.log(id)
                var xmlhttp = new XMLHttpRequest();
                if (favoriteBtn.classList.contains("active")) {
                    const favorite = 0
                    xmlhttp.open("GET", "/fashionShop/controllers/updateFavorite.asp?q=" + favorite +"&id="+id, true);
                    xmlhttp.send();
                } else {
                    const favorite = 1
                    xmlhttp.open("GET", "/fashionShop/controllers/updateFavorite.asp?q=" + favorite +"&id="+id, true);
                    xmlhttp.send();
                }
            })
        });
    </script>
    <!-- Popper js -->
    <script src="js/popper.min.js"></script>
    <!-- Bootstrap js -->
    <script src="js/bootstrap.min.js"></script>
    <!-- Plugins js -->
    <script src="js/plugins.js"></script>
    <!-- Classy Nav js -->
    <script src="js/classy-nav.min.js"></script>
    <!-- Active js -->
    <script src="js/active.js"></script>

</body>

</html>