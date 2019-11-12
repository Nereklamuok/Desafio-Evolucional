<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="Desafio_Evolucional.WebForm2" %>

<!DOCTYPE html>
<link href="Default.css" rel="stylesheet" type="text/css" />
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <ul>
                <li>
                    <asp:Label ID="Label1" runat="server" Text="Gerar SQL: "></asp:Label>
                    <asp:Button ID="InsertEntriesBtn" runat="server" Text="SQL" OnClick="InsertEntriesBtn_Click" />
                </li>
                <li>
                    <asp:Label ID="Label2" runat="server" Text="Gerar Excel:"></asp:Label>
                    <asp:Button ID="GenerateExcelBtn" runat="server" Text="Excel" OnClick="GenerateExcelBtn_Click" />
                </li>
                <li>
                    <asp:Button ID="ButtonLogout" runat="server" Text="Sair" OnClick="ButtonLogout_Click" />
                    <asp:Button ID="ButtonDownload" runat="server" Text="Download" OnClick="ButtonDownload_Click" />
                </li>
            </ul>
        </div>
    </form>
</body>
</html>
