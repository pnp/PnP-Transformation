<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Thanks.aspx.cs" Inherits="EmployeeRegistration.FormsWeb.Pages.Thanks" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div><h1>Thank you!</h1></div>
    <div>
       Registered employee successfully
    </div>
        <p>
            <asp:LinkButton ID="lnkNewAppPage" Text="Add new employee" OnClick="lnkNewAppPage_Click" runat="server"></asp:LinkButton>
            
        </p>
    </form>
</body>
</html>
