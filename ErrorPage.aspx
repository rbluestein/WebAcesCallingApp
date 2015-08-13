<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ErrorPage.aspx.vb" Inherits="ErrorPage" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title runat="server" id="PageCaption">Error</title>    
    <script src="jscripts/backfix.min.js" type="text/javascript"></script>    
    <link rel="Stylesheet" type="text/css" href="skin/enroller_login.css" />

   <script type="text/javascript" >
        bajb_backdetect.OnBack = function() {
            window.history.go(1);
        }
    </script>    
    
    
	
</head>
<body>
<form>

    <div id="enroller_login_error">
        <h1>Oops...</h1>
        <h3>An error has occured.</h3>
        <p><asp:Literal ID="litMessage" runat="server"></asp:Literal></p>
        <div class="clear"></div>
        <a class="button close" onclick="javascript:fnCloseWindow()"><span>Close</span></a>
    </div>
    
    <script type="text/javascript" >
        function fnCloseWindow() {
            if (navigator.appName.indexOf("Microsoft") > -1) {
                window.open('','_self',''); window.close()
            } else {
                var w=window.open('about:blank','_self');w.close()         
            }      
        }           
    </script>
    </form>
</body>
</html>
