<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="EmployeeRegistration.FormsWeb.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

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
        <div>
            <table style="left: 50px; top: 50px; position: absolute;">
                <tr>
                    <td colspan="2">
                        <h1>Employee Registration Form</h1>
                    </td>
                </tr>
                <tr>
                    <td>Emp Number</td>
                    <td>
                        <asp:TextBox ID="txtEmpNumber" runat="server"></asp:TextBox></td>
                </tr>
                <tr>
                    <td>User ID</td>
                    <td>
                        <asp:TextBox ID="txtUserID" runat="server"></asp:TextBox>
                        <asp:Button ID="btnGetManager" Text="Get name and manager from profile" OnClick="btnGetManager_Click" runat="server" />
                    </td>
                </tr>
                <tr>
                    <td>Name</td>
                    <td>
                        <asp:TextBox ID="txtName" runat="server"></asp:TextBox>
                    </td>
                </tr>
                <tr>
                    <td>Manager</td>
                    <td>
                        <asp:TextBox ID="txtManager" runat="server"></asp:TextBox></td>
                </tr>
                <tr>
                    <td>Designation</td>
                    <td>
                        <asp:DropDownList ID="ddlDesignation" DataTextField="Designation" DataValueField="Designation" runat="server"></asp:DropDownList></td>
                </tr>

                <tr>
                    <td>Location</td>
                    <td>
                        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
                            <ContentTemplate>
                                <asp:DropDownList ID="ddlCountry" DataTextField="Country" DataValueField="Country"
                                    OnSelectedIndexChanged="ddlCountry_SelectedIndexChanged" AutoPostBack="true" runat="server">
                                </asp:DropDownList>
                                <asp:DropDownList ID="ddlState" DataTextField="State" DataValueField="State" OnSelectedIndexChanged="ddlState_SelectedIndexChanged" AutoPostBack="true" runat="server"></asp:DropDownList>
                                <asp:DropDownList ID="ddlCity" DataTextField="City" DataValueField="City" runat="server"></asp:DropDownList>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top">Skills</td>
                    <td>
                        <asp:UpdatePanel ID="upPanael" runat="server">
                            <ContentTemplate>
                                <table>
                                    <tr>
                                        <th>Technology
                                        </th>
                                        <th>Experience (in months)
                                        </th>
                                    </tr>
                                    <asp:Repeater ID="rptSkills" runat="server">
                                        <ItemTemplate>
                                            <tr>
                                                <td>
                                                    <asp:TextBox ID="rptTxtTechnology" runat="server"
                                                        Width="150px" MaxLength="50" Text='<%#Eval("Technology") %>'>
                                                    </asp:TextBox>
                                                </td>
                                                <td>
                                                    <asp:TextBox ID="rptTxtExperience" runat="server"
                                                        Width="150px" MaxLength="50" Text='<%#Eval("Experience") %>'>
                                                    </asp:TextBox>
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                    </asp:Repeater>
                                    <tr>
                                        <td colspan="2">
                                            <asp:LinkButton ID="lnkAddSkills" OnClick="lnkAddSkills_Click" Text="Add New Skill" runat="server"></asp:LinkButton>
                                        </td>
                                    </tr>
                                </table>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                </tr>
                <tr>
                    <td style="vertical-align: top">Attachments</td>
                    <td>
                        <input id="hdnAttachmentID" type="hidden" runat="server" />

                        <table>
                            <asp:UpdatePanel ID="upAttachments" runat="server">
                                <ContentTemplate>
                                    <asp:Repeater ID="rptUploadedFiles" runat="server">
                                        <ItemTemplate>
                                            <tr>
                                                <td>
                                                    <asp:HyperLink Target="_blank" ID="rptAttachment" runat="server"
                                                        Width="150px" MaxLength="50" Text='<%#Eval("FileName") %>' NavigateUrl='<%#Eval("FileUrl") %>'>
                                                    </asp:HyperLink>
                                                </td>
                                                <td>
                                                    <asp:LinkButton ID="rptDelete" Text="Delete" CommandArgument='<%#Eval("FileRelativeUrl") %>' OnClick="rptDelete_Click" runat="server">
                                                    </asp:LinkButton>
                                                </td>
                                            </tr>
                                        </ItemTemplate>
                                    </asp:Repeater>
                                </ContentTemplate>
                            </asp:UpdatePanel>
                            <tr>
                                <td>
                                    <asp:FileUpload ID="empAttachment" runat="server" />
                                </td>
                                <td>
                                    <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="btnUpload_Click" />
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:Label ID="lblFileUrl" runat="server"></asp:Label>

                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <!-- ToDo: Add user to site group --> 
                <tr>
                    <td>

                        <asp:Button ID="btnSave" Text="Save" Visible="false" OnClick="btnSave_Click" runat="server" />
                        <asp:Button ID="btnUpdate" Text="Update" Visible="false" OnClick="btnUpdate_Click" runat="server" />
                    </td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
