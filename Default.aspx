<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <script type="text/javascript">
        function CheckFileCompare() {
            if (document.getElementById("File1").value == "" || document.getElementById("File2").value == "") {
                alert("Please choose File for compare! ");
                return false;
            } else {
                if (confirm("Are you sure close excel") == true) {
                    return true;
                } else {
                    return false;
                }
            }
        }
    </script>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
            <table style="width: 100%; border: 0px;" border="0">
                <tr>
                    <td colspan="3">&nbsp;<br />
                        &nbsp;</td>
                </tr>
                <tr>
                    <td style="width: 30%"></td>
                    <td style="width: 40%">
                        <table style="width: 100%; border-style: inset; border-width: 2px; border-color: aqua;" cellpadding="5">
                            <tr>
                                <td colspan="2" style="text-align: center">
                                    <br />
                                    <b>COMPARE EXCEL</b></td>
                            </tr>
                            <tr>
                                <td>File 1: </td>
                                <td>
                                    <asp:FileUpload ID="File1" runat="server" Width="100%" /></td>
                            </tr>
                            <tr>
                                <td>File 2: </td>
                                <td>
                                    <asp:FileUpload ID="File2" runat="server" Width="100%" /></td>
                            </tr>
                            <tr>
                                <td>&nbsp;</td>
                                <td>
                                    <asp:Button ID="btnCompare" runat="server" Text="COMPARE" OnClientClick="return CheckFileCompare();" /><br />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">&nbsp;</td>
                            </tr>
                        </table>
                    </td>
                    <td style="width: 30%"></td>
                </tr>
            </table>
        </div>
    </form>
</body>
</html>
