<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>WebAces Enroller Login</title>
    <link rel="Stylesheet" type="text/css" href="skin/enroller_login.css" />
</head>
<body onload="DoRedirect()">
    <form id="form1" runat="server">
    <%--<div>--%>

<asp:Panel ID="pnlLogin" runat="server" Visible="False">

<%--        <script type="text/javascript">
            function DoRedirect() {var i}
        </script>--%>
                        <div id="login_trow">
                            <div id="login_left"></div>
       
                            <div id="login_center">
                                <h1 id="login_heading">WebAces Enroller Login</h1>
                                <div id="clientname">
                                    for<span>&nbsp;<asp:Literal ID="litClientName" runat="server"></asp:Literal></span>  
                                </div>                        
                                    <%--<div class="clear"></div>--%>                                                                                                
                                <table>
                                    <colgroup>
                                        <col class="login_labels" />
                                        <col class="login_fields" />
                                    </colgroup>
                                    <tr>
                                        <td>
                                            <label for="username">Enroller ID:</label>
                                        </td>
                                        <td><div class="login_field">
                                            <div class="input_field">
                                                <asp:TextBox ID="txtLoginUserName" runat="server" MaxLength="20" TabIndex="10"></asp:TextBox>
                                            </div>
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <label for="password">Password:</label>
                                        </td>
                                        <td><div class="login_field">
                                            <div class="input_field">
                                                <asp:TextBox ID="txtLoginPassword" runat="server" TextMode="Password" 
                                                    MaxLength="50" TabIndex="11"></asp:TextBox>
                                            </div>
                                            </div>
                                        </td>
                                    </tr>                                                                                           
                                    <tr>                                    
                                        <td colspan="2">
                                           <div class="login_buttons">                                            
                                                <input id="btnSubmit" type="button" value="Login" onclick="fnSubmit()" />
                                                <div class="clear"></div>
                                            </div>                               
                                         </td>
                                    </tr>
                                </table>
                                <div>
                                    <asp:Label ID="lblLoginMessage" runat="server"></asp:Label>
                                </div> 
                                <div class="clear"></div>
                            </div>  
                            
                            <div id="login_right"></div>                    
                        </div>
                        <div id="enroller_login_bottom"></div>
                </asp:Panel>

    
        <asp:Literal ID="litRedirect" runat="server" EnableViewState="False"></asp:Literal>
    
    <%--</div>--%>&nbsp;<asp:Literal ID="litBrowserName" runat="server" 
        EnableViewState="False"></asp:Literal>
    </form>
    
    <script type="text/javascript">
        var InProgress = 0
        function fnSubmit() {
            if (InProgress == 1) {
                return
            } else {
                InProgress = 1
                document.getElementsByTagName("form")[0].submit()
            }
        }


        function EnterKeySubmit() {
            var BrowserName = document.getElementById("hdBrowserName").value
            var KeyID

            if (BrowserName == "IE") {
                KeyID = window.event.keyCode
            } else {
                KeyID = e.charCode
            }
            if (KeyID == 13) {
                fnSubmit()
            }
        }            
    </script>
    
</body>
</html>
