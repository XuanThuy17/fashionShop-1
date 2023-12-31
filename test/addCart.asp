
<%
'     Set id = Request.QueryString("id")
'     set mycarts = Server.CreateObject("Scripting.Dictionary")
'     mycarts.add "cart1", 1
'     mycarts.add "cart2", 2
'     mycarts.add "cart3", 3
'     Response.ContentType = "application/json"
'     ' Response.Write "{""messenger"": ""Product is not exists your cart."", "
'     ' Response.Write """totalProduct"": ""total"&id&"""}"

'     Response.Write "["
'     for i = 1 to 5
'     Response.Write    "{ ""id"": "&i&", ""value"": "&i&"},"
'     ' Response.Write    "{ ""id"": 2, ""value"": 2 },"
'     next
'     Response.Write    "{ ""id"": 6, ""value"": 6 }"
'     Response.Write "]"

%>

<%
Set Conn = Server.CreateObject("ADODB.Connection")
Conn.Open "Provider=SQLOLEDB.1;Data Source=THUY092\SQLEXPRESS;Database=shop;User Id=sa;Password=123"

sql = "select cart.ID_product, product.name, cart.ID_size, size.size, cart.ID_color, color.color, brand, product.price, cart.quantity, link1, sale_percent, end_day from cart inner join product_size_color p on cart.ID_product = p.ID_product inner join size on cart.ID_size = size.ID_size inner join color on cart.ID_color = color.ID_color inner join product on product.ID_product = cart.ID_product inner join brand on brand.ID_product = cart.ID_product inner join imageProduct on imageProduct.ID_product = cart.ID_product inner join discount on discount.ID_product = cart.ID_product where cart.ID_user = 1 group by cart.ID_product, product.name, cart.ID_size, size.size, cart.ID_color, color.color, brand, product.price, cart.quantity, link1, sale_percent, end_day"

Set result = Conn.execute(sql)

Response.ContentType = "application/json"
Response.Write "["
do while not result.EOF
    price = CInt(result("price"))
    quantity = CInt(result("quantity"))
    percent = CInt(result("sale_percent"))

    ' tạo biến currentDate lấy ra ngày hiện tại
    ' so sánh với ngày hết hạn sale để tính giá tiền
    currentDate_cart = Date()

    datee = FormatDateTime(result("end_day"),2)

    if (CStr(datee) < CStr(currentDate_cart)) then
        percent = 0
    end if

    Response.Write "{ ""id"": """&result("ID_product")&""", ""id_size"": """&result("ID_size")&""", ""id_color"": """&result("ID_color")&""", ""name"": """&result("name")&""", ""size"": """&result("size")&""", ""color"": """&result("color")&""", ""brand"": """&result("brand")&""", ""price"": """&result("price")&""", ""quantity"": """&result("quantity")&""", ""sale_percent"": """&percent&""", ""end_day"": """&result("end_day")&""", ""link1"": """&result("link1")&"""},"
        
result.MoveNext
loop
Response.Write    "{ ""id"": ""-1"", ""name"": ""6"" }"
Response.Write "]"
%>
