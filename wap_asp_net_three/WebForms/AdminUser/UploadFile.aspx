<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="UploadFile.aspx.cs" Inherits="wap_asp_net_three.WebForms.AdminUser.UploadFile" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.13.5/xlsx.full.min.js"></script>
    <script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.13.5/jszip.js"></script>
    <script src="js/UploadFile.js?1500"></script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <input type="file" id="fileUpload" />
            <input type="button" id="upload" value="Upload" onclick="Upload()" />
            <hr />
            <div id="div_results"></div>
            <div id="dvExcel"></div>
        </div>
    </form>
</body>
</html>
