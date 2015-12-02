<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Thanks.aspx.cs" Inherits="EmployeeRegistration.FormsWeb.Pages.Thanks" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript" src="../Scripts/jquery-1.9.1.js"></script>
    <script type="text/javascript" src="../Scripts/app.js"></script>
</head>
<body style="display: none; overflow: auto;">
    <form id="form1" runat="server">
        <asp:ScriptManager ID="scriptManager" runat="server" EnableCdn="True"></asp:ScriptManager>
        <div id="divSPChrome"></div>
        <div style="left: 50px; position: absolute;">
            <div>
                <h1>Thank you!</h1>
            </div>
            <div>
                Registered employee successfully
            </div>
            <p>
                <asp:LinkButton ID="lnkNewAppPage" Text="Add new employee" OnClick="lnkNewAppPage_Click" runat="server"></asp:LinkButton>
            </p>
        </div>
    </form>
</body>
</html>
