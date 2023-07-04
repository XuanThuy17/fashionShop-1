<% 'code here
Dim connDB
set connDB = Server.CreateObject("ADODB.Connection")
Dim strConnection
strConnection = "Provider=SQLOLEDB.1;Data Source=THUY092\SQLEXPRESS;Database=shop;User Id=sa;Password=123"
connDB.ConnectionString = strConnection
connDB.Open()

ID_employeeEdit = Request.QueryString("id")
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
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" integrity="sha512-iecdLmaskl7CVkqkXNQ/ZH/XLlvWZOJyj7Yy7tcenmpD1ypASozpmT/E0iPtmFIB46ZmdtAc9eNBvH0H/ZpiBw==" crossorigin="anonymous" referrerpolicy="no-referrer" />
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
  .file-upload-info {
    padding-bottom: 25px;
  }
  .pass-error {
    display: flex;
    flex-direction: column;
    padding-left: 1em;
  }
  .pass-error span {
    font-size: 14px;
  }
  .pass-error .pass-suc {
    color: #00cc1f;
  }
  .pass-error .pass-err {
    color: #ff0700;
  }
  .imgUp {
      width: 5em;
      margin-right: 1em;
      object-fit: cover;
  }
</style>
<body>
  <div class="container-scroller"> 
    <!-- partial:partials/_navbar.html -->
    
    <!-- #include file="../../partials/_header.asp" --> 

    <!-- partial -->
    <div class="container-fluid page-body-wrapper">
      <!-- #include file="../../partials/_settings-panel.asp" -->

      <!-- #include file="../../partials/_sidebar.asp" -->
      
      <!-- partial -->
      <div class="main-panel">
        <div class="content-wrapper">
          <div class="row">
            <div class="col-12 grid-margin stretch-card">
              <div class="card">
                <div class="card-body">
                    <h4 class="card-title">Edit Employee</h4>
                    <p class="card-description">
                        Edit Employee
                    </p>
                    <%
                    Dim firstName, lastName, cmnd, phoneNumber, birthDay, joinDate, gender, avatar
                    Dim sql, rs

                    ' Lấy dữ liệu từ biểu mẫu HTML
                    firstName = Request.Form("firstName")
                    lastName = Request.Form("lastName")
                    cmnd = Request.Form("cmnd")
                    phoneNumber = Request.Form("phone_number")
                    birthDay = Request.Form("birthDay")
                    joinDate = Request.Form("join")
                    gender = Request.Form("gender")
                    avatar = Request.Form("avt")

                    ' Mở kết nối với cơ sở dữ liệu
                    connDB.Open()


                    ' Tạo chuỗi SQL để cập nhật thông tin người dùng
                    sql = "UPDATE Users SET firstName='" & firstName & "', lastName='" & lastName & "', cmnd='" & cmnd & "', phone_number='" & phoneNumber & "', birthday='" & birthDay & "', joindate='" & joinDate & "', gender='" & gender & "', avatar='" & avatar & "' WHERE userID=" & ID_user

                    ' Thực hiện câu truy vấn SQL
                    connDB.Execute(sql)

                    ' Đóng kết nối cơ sở dữ liệu và chuyển hướng người dùng đến trang thành công
                    connDB.Close()
                    <!-- Response.Redirect "../allemployee.asp" -->
                    %>
                    <form class="forms-sample">
                        <div class="form-group">
                        <label for="exampleInputName1">First Name</label>
                        <input name="firstName" type="text" value="<%=Result("firstName")%>" class="form-control" id="exampleInputName1" required placeholder="First Name">
                        </div>
                        <div class="form-group">
                        <label for="exampleInputName1">Last Name</label>
                        <input name="lastName" type="text" value="<%=Result("lastName")%>" class="form-control" id="exampleInputName1" required placeholder="Last Name">
                        </div>
                        <div class="form-group">
                        <label for="exampleInputName1">Identity Card</label>
                        <input name="cmnd" type="text" value="<%=Result("cmnd")%>" class="form-control" id="exampleInputName1" required placeholder="Identity Card">
                        </div>
                        <div class="form-group">
                        <label for="exampleInputCity1">Phone Number</label>
                        <input name="phone_number" type="text" value="<%=Result("phone_number")%>" class="form-control" id="exampleInputCity1" required placeholder="Phone Number">
                        </div>
                        <div class="form-group">
                        <label for="exampleInputCity1">Birthday</label>
                        <input name="birthDay" type="date" value="<%=Result("birthday")%>" class="form-control" id="birthdayID" required placeholder="Birthday">
                        </div>
                        <div class="form-group">
                        <label for="exampleInputCity1">Join on</label>
                        <input name="join" type="date" value="<%=Result("joindate")%>" class="form-control" id="exampleInputCity1" required placeholder="Official Working Day">
                        </div>
                        <div class="form-group">
                        <label for="exampleSelectGender">Gender</label>
                            <select name="gender" class="form-control" id="selectGender" required>
                                <% if (Result("gender")) then %>
                                    <option value="1">Male</option>
                                    <option value="2">Female</option>
                                <% else %>
                                    <option value="2">Female</option>
                                    <option value="1">Male</option>
                                <% end if%>
                            </select>
                        </div>
                        <div class="form-group">
                        <label>File upload</label>
                        <input type="file" name="avt" class="file-upload-default">
                        <div class="input-group col-xs-12">
                            <input type="file" class="form-control file-upload-info" required placeholder="Upload Image" onchange="handleFileUpload(this)">
                        </div>
                        </div>
                        <ol id="filelist" style="display: flex;">

                        </ol>	
                        <input class="nameAvt" type="hidden" value="<%=Result("avatar")%>">
                        
                        <button type="submit" class="submit btn btn-primary me-2">Submit</button>
                        <button class="cancel btn btn-light">Cancel</button>
                    </form>
                    <%
                    Result.MoveNext
                    loop
                    %>
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
    
    </div>
    <!-- #include file="../../js/mainJs.asp" -->
</body>

</html>