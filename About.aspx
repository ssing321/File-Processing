<%@ Page Title="About" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="About.aspx.cs" Inherits="ProjExcel.About" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">


    <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Size="XX-Large" Text="Excel to GridView "></asp:Label>
    <br />
    <br />
    <asp:Label ID="Label2" runat="server" Font-Size="Large" ForeColor="#9966FF" Text="Upload the Excel File and read data using gridview"></asp:Label>
    <br />
    <br />
    <table border="1" class="nav-justified" style="height: 134px; width: 35%; background-color: #00CCFF">
        <tr>
            <td style="height: 68px; width: 282px; background-color: #FF99FF">Excel File here </td>
            <td style="height: 68px">
                <asp:FileUpload ID="FileUpload1" runat="server" />
            </td>
        </tr>
        <tr>
            <td class="modal-sm" style="width: 282px; background-color: #FF99FF">
                <br />
            </td>
            <td>
                <asp:Button ID="Button1" runat="server" style="background-color: #FFFFFF" Text="Upload" OnClick="Button1_Click" />
                <br style="background-color: #FF99FF" />
                <asp:Label ID="Label3" runat="server" style="background-color: #FFFFFF"></asp:Label>
                <br />
                <br />
                <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="UpldDatabase" />
                <br />
                <asp:Label ID="Label4" runat="server" Text="Label"></asp:Label>
            </td>
        </tr>
    </table>
    <br />


    <asp:GridView ID="GridView1" runat="server" BackColor="White" BorderColor="#336666" BorderStyle="Double" BorderWidth="3px" CellPadding="4" GridLines="Horizontal" Height="90px" Width="576px">
        <FooterStyle BackColor="White" ForeColor="#333333" />
        <HeaderStyle BackColor="#336666" Font-Bold="True" ForeColor="White" />
        <PagerStyle BackColor="#336666" ForeColor="White" HorizontalAlign="Center" />
        <RowStyle BackColor="White" ForeColor="#333333" />
        <SelectedRowStyle BackColor="#339966" Font-Bold="True" ForeColor="White" />
        <SortedAscendingCellStyle BackColor="#F7F7F7" />
        <SortedAscendingHeaderStyle BackColor="#487575" />
        <SortedDescendingCellStyle BackColor="#E5E5E5" />
        <SortedDescendingHeaderStyle BackColor="#275353" />
    </asp:GridView>


</asp:Content>
