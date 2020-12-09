<!--#include file ="..\ayar.asp"-->
<%
parola = request.form("parola" )
sayfa = request.querystring("sayfa")
if len(sayfa) < 1 then
sayfa="default.asp"
end if
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Giriş</title>
</head>

<body>
<%
If parola = yonetimsifre Then
Session.Timeout=10
Session("siteyoneticisi")= True
end if
If Session("siteyoneticisi") then
%>
<meta http-equiv="refresh" content="1;url=<%=sayfa%>">
<%
else
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="760" id="AutoNumber1">
  <tr>
    <td width="33%">&nbsp;</td>
    <td width="33%">
    <form method="POST" action="giris.asp?sayfa=<%=sayfa%>"> 

  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family:Arial Black; font-size:8pt" width="100%" >
    <tr>
      <td align="center">
      <p align="center">ŞİFRE</td>
    </tr>
    <tr>
      <td align="center">
      <p align="center"><input type="password" name="parola" size="20"></td>
    </tr>
    <tr>
      <td align="center">
      <p align="center"><input type="submit" value="Gir" name="admin"></td>
    </tr>
  </table> 
</form> 
</td>
    <td width="34%">&nbsp;</td>
  </tr>
</table>
<%end if%>
</body>

</html>