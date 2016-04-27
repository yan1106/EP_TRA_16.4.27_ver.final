<%@ Page Language="C#" AutoEventWireup="true" CodeFile="EP_Category_Add.aspx.cs" Inherits="EP_Category_Add" %>

<!DOCTYPE html>

   <html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>PreNPI Import Excel Data </title>
      <link rel="stylesheet" href="..\css\styles.css" type="text/css" />
      <link rel="stylesheet" href="..\css\styles.css" type="text/css" />
      <link rel="stylesheet" href="http://code.jquery.com/ui/1.9.2/themes/base/jquery-ui.css" />
      <script src="http://code.jquery.com/jquery-1.8.3.js"></script>
      <script src="http://jqueryui.com/resources/demos/external/jquery.bgiframe-2.1.2.js"></script>
      <script src="http://code.jquery.com/ui/1.9.2/jquery-ui.js"></script>
      <script type="text/javascript">

</script>
    <style type="text/css">
        .style1
        {
            font-size: large;
        }
        .style2
        {
            height: 20px;
        }
    </style>
</head>
<body>
    <form id="form2" runat="server">
    <div>
     <fieldset style="margin:auto;" class="fieldset">
    <legend class="legend" style="font-weight: 700; font-size: large;">Import Excel&nbsp; Category_Data</legend>
        <table style="width:100%;">           
            <tr>
                <td>
                    <strong><span class="style1">匯入Excel時,請您注意下列事項:</span></strong><br />
                    1.Excel檔案類型必須是[.xlsx].<br />
                    2.檔名名稱不影響存放資料的結果<br />
                    3.Stage的名稱是以工作表的名稱為命名<br />
                    4.工作表裡共有五欄位，名稱依序為<span style="color:blue;">(Effect stage,Input Items,Special Characteristics,Methodology,Category,Key parameter)</span><br />
                    5.範例下載:<asp:HyperLink ID="HyperLink1" NavigateUrl="stage.xlsx" runat="server">Category_Sample</asp:HyperLink><br />

                </td>
            </tr>
            <tr>
                <td class="style2">              
                </td>
            </tr>
            <tr>
                <td>
    <asp:FileUpload runat="server" ID="FileUploadToServer" Width="382px" Height="36px">
        </asp:FileUpload>
        <asp:Button runat="server" Text="Import" ID="btnUpload" onclick="btnUpload_Click" 
                        Height="25px" class="blueButton button2" />
                </td>
            </tr>
            <tr>
                <td>
        <asp:Label ID="lblMsg" runat="server" ForeColor="Green" Text=""></asp:Label>        
                </td>
            </tr>
            <tr>
                <td>
                <hr/>
           <asp:GridView runat="server" ID="gvRecord" EmptyDataText="No record found!"
            Height="25px" BackColor="White" BorderColor="#CCCCCC" BorderStyle="None" 
            BorderWidth="1px" CellPadding="3">
               <FooterStyle BackColor="White" ForeColor="#000066" />
               <HeaderStyle BackColor="#006699" Font-Bold="True" ForeColor="White" />
               <PagerStyle BackColor="White" ForeColor="#000066" HorizontalAlign="Left" />
               <RowStyle ForeColor="#000066" />
               <SelectedRowStyle BackColor="#669999" Font-Bold="True" ForeColor="White" />
               <sortedascendingcellstyle backcolor="#F1F1F1" />
               <sortedascendingheaderstyle backcolor="#007DBB" />
               <sorteddescendingcellstyle backcolor="#CAC9C9" />
               <sorteddescendingheaderstyle backcolor="#00547E" />
        </asp:GridView>
                </td>
            </tr>
        </table>
         </fieldset>    
  
    </form>
</body>
</html>

  