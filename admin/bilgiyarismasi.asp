<%
If Session("siteyoneticisi") then
%>
<!--#include file ="../ayar.asp"-->
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Bilgi Yarışması</title>
</head>
<body>
<table border="1" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" id="AutoNumber1" cellpadding="0" width="750">
  <tr>
    <td width="748" align="center" colspan="2">
    <nobr><b><font size="2"><a href="bilgiyarismasi.asp?bolum=onaylisorular"><font color="#000000">Onaylı Sorular</font></a>(<%for w=1 to 25%><a href="bilgiyarismasi.asp?bolum=onaylisorular&seviye=<%=w%>"><font color="#000000"><%=w%></font></a>-<%next%>)</font></b></nobr>&nbsp;
    <nobr><b><font size="2"><a href="bilgiyarismasi.asp?bolum=onaysizsorular"><font color="#000000">Onaysız Sorular</font></a></font></b></nobr>&nbsp;
    <nobr><b><font size="2"><a href="bilgiyarismasi.asp?bolum=soruekle"><font color="#000000">Soru Ekle</font></a></font></b></nobr>&nbsp;
    <nobr><b><font size="2"><a href="cik.asp"><font color="#000000">Çıkış</font></a></font></b></nobr>&nbsp;
    </td>
  </tr>
</table>
<%
bolum=request.querystring("bolum")
if len(Request.querystring("no"))>0 then
no = cint(Request.querystring("no"))
end if
if len(Request.querystring("sayfa"))>0 then
sayfa = Request.querystring("sayfa")
else
sayfa = "1"
end if

Set bag = Server.CreateObject("ADODB.Connection" ) 
bag.Open ("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & veritabaniyonetim)  
%>
<%
Select Case bolum
%>

<%
	Case "onaylisorular"
%>
<%
if len(request.querystring("seviye"))>0 then
eksql=" and sorular.derece="&request.querystring("seviye")
else
eksql=""
end if
set rs = server.createobject("ADODB.Recordset")
rs.open "SELECT sorular.* FROM sorular where sorular.onay=true"&eksql,bag,1,3
if not rs.eof then
rs.pagesize = 10
rs.absolutepage = sayfa
tsayfa = rs.pagecount
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family:Verdana; font-size:10pt; font-weight:bold" bordercolor="#111111" width="750">
<%
for i=1 to rs.pagesize
if rs.eof then exit for 
%>
  <tr>
    <td colspan="3" width="100%" valign="middle" align="center">
Göderen : <%=rs("ekleyen")%><br>
Tür : <%=rs("grup")%> - Seviye : <%=rs("derece")%><br>
Cevaplanma Sayısı : <%=rs("csayisi")%> - Bilinme Sayısı : <%=rs("bsayisi")%> - Bilinme Oranı : <%if rs("csayisi")>0 then response.write("%"&int((rs("bsayisi")/rs("csayisi"))*100)) end if%><br>
Soru : <%=rs("soru")%><br>
a)<%=rs("a")%> b)<%=rs("b")%> c)<%=rs("c")%> d)<%=rs("d")%><br>
Doğru Cevap : <%=rs("dogru")%>
    </td>
  </tr>
  <tr>
    <td width="33%" align="center"><a href="bilgiyarismasi.asp?bolum=soruduzenle&no=<%=rs("no")%>"><font color="#000000">Düzelt</font></a></td>
    <td width="33%" align="center"><a href="bilgiyarismasi.asp?bolum=sorusil&no=<%=rs("no")%>"><font color="#000000">Sil</font></a></td>
    <td width="33%" align="center">
<form method="POST" action="bilgiyarismasi.asp?bolum=sartir&no=<%=rs("no")%>" style="margin-top: 0; margin-bottom: 0">
<select name="seviye" onchange="this.form.submit()" style="font-size: 10pt; font-family: Verdana; font-weight: bold">
<option selected>Seviye</option>
<%
for x=1 to 25
%>
<option value="<%=x%>"><%=x%></option>
<%
next
%>
</select>
</form>
</td>
  </tr>
<%
rs.movenext
next
end if
usayisi=rs.recordcount
rs.Close
set rs=nothing
%> 
</table>
<%
for y=1 to tsayfa 
if CINT(TRIM(sayfa))=CINT(TRIM(y)) then%>
<b>[<%=y%>]</b>
<%
else
if len(request.querystring("seviye"))>0 then
ek="&seviye="&request.querystring("seviye")
end if
response.write "<a href='bilgiyarismasi.asp?bolum=onaylisorular&sayfa=" & y & ek&"'>" & y & "</a>&nbsp;"
end if
%><%next%>
<br><%if len(request.querystring("seviye"))>0 then
response.write request.querystring("seviye")&". Seviye"
else%>
Toplam<%end if%> Onaylı Soru Sayısı : <%=usayisi%>
<%
	Case "onaysizsorular"
%>
<%
set rs = server.createobject("ADODB.Recordset")
rs.open "SELECT sorular.* FROM sorular where sorular.onay=false",bag,1,3
if not rs.eof then
rs.pagesize = 10
rs.absolutepage = sayfa
tsayfa = rs.pagecount
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family:Verdana; font-size:10pt; font-weight:bold" bordercolor="#111111" width="750">
<%
for i=1 to rs.pagesize
if rs.eof then exit for
%>
  <tr>
    <td colspan="4" width="100%" valign="middle" align="center">
Göderen : <%=rs("ekleyen")%><br>
Tür : <%=rs("grup")%> - Seviye : <%=rs("derece")%><br>
Soru : <%=rs("soru")%><br>
a)<%=rs("a")%> b)<%=rs("b")%> c)<%=rs("c")%> d)<%=rs("d")%><br>
Doğru Cevap : <%=rs("dogru")%>
    </td>
  </tr>
  <tr>
    <td width="25%" align="center"><a href="bilgiyarismasi.asp?bolum=soruonayla&no=<%=rs("no")%>"><font color="#000000">Onayla</font></a></td>
    <td width="25%" align="center"><a href="bilgiyarismasi.asp?bolum=soruduzenle&no=<%=rs("no")%>"><font color="#000000">Düzelt</font></a></td>
    <td width="25%" align="center"><a href="bilgiyarismasi.asp?bolum=sorusil&no=<%=rs("no")%>"><font color="#000000">Sil</font></a></td>
    <td width="25%" align="center">
<form method="POST" action="bilgiyarismasi.asp?bolum=sartir&no=<%=rs("no")%>" style="margin-top: 0; margin-bottom: 0">
<select name="seviye" onchange="this.form.submit()" style="font-size: 10pt; font-family: Verdana; font-weight: bold">
<option selected>Seviye</option>
<%
for x=1 to 25
%>
<option value="<%=x%>"><%=x%></option>
<%
next
%>
</select>
</form>
</td>
  </tr>
<%
rs.movenext
next
end if
usayisi=rs.recordcount
rs.Close
set rs=nothing
%> 
</table>
<%
for y=1 to tsayfa 
if CINT(TRIM(sayfa))=CINT(TRIM(y)) then%>
<b>[<%=y%>]</b>
<%
else
response.write "<a href='bilgiyarismasi.asp?bolum=onaysizsorular&sayfa=" & y & "'>" & y & "</a>&nbsp;"
end if
%><%next%><br>Toplam Onaysız Soru Sayısı : <%=usayisi%>
<%
	Case "soruekle"
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family:Verdana; font-size:10pt; font-weight:bold" bordercolor="#111111" width="750">
<tr>
<td width="100%" valign="middle" align="center">
<form method="POST" style="margin-top: 0; margin-bottom: 0;" action="bilgiyarismasi.asp?bolum=sorukaydet">
              <label><strong>Tür :
              <input type="text" name="tur" />
              </strong></label>
              <strong><br />
              <label>Soru :<br />
              <textarea name="soru" cols="50" rows="4"></textarea>
              </label>
              <br />
              <label>A :
              <input type="text" name="a" />
              <input name="dogru" type="radio" value="a" />
              </label>
              <br />
              <label>B :
              <input type="text" name="b" />
              <input name="dogru" type="radio" value="b" />
              </label>
              <br />
              <label>C :
              <input type="text" name="c" />
              <input name="dogru" type="radio" value="c" />
              </label>
              <br />
              <label>D :
              <input type="text" name="d" />
              <input name="dogru" type="radio" value="d" />
              </label><br>
              <label>Seviye :
		<select name="seviye">
		<option>---</option>
		<%
		for b=1 to 25
		%>
		<option value ="<%=b%>"><%=b%></option>
		<%
		next
		%>
		</select>
              </label>
              <br />
		<input type="submit" value="Gönder" />
                </strong>
            </form>
</td>
</tr>
</table>

<%
	Case "sorukaydet"
%>
<%
		Function avla(byval hedef) 
		hedef = replace(hedef,"_","") 
		hedef = replace(hedef,"*","")  
		hedef = replace(hedef,"%","") 
		hedef = replace(hedef,"<","") 
		hedef = replace(hedef,">","") 
		hedef = replace(hedef,"chr(13)","<br>") 
		avla=trim(hedef) 
		End Function 
		if len(request.form("tur"))>0 and len(request.form("soru"))>0 and len(request.form("seviye"))>0 and len(request.form("a"))>0 and len(request.form("b"))>0 and len(request.form("c"))>0 and len(request.form("d"))>0 and len(request.form("dogru"))>0 then
		set rs = server.createobject("ADODB.Recordset")
		rs.open "select sorular.* from sorular",bag,1,3
		rs.addnew
		rs("soru")=avla(request.form("soru"))
		rs("grup")=avla(request.form("tur"))
		rs("a")=avla(request.form("a"))
		rs("b")=avla(request.form("b"))
		rs("c")=avla(request.form("c"))
		rs("d")=avla(request.form("d"))
		rs("dogru")=avla(request.form(request.form("dogru")))
		rs("ekleyen")="Yönetim"
		rs("onay")=true
		rs("bsayisi")=0
		rs("derece")=request.form("seviye")
		rs("csayisi")=0
		rs.update
		rs.close
		set rs = nothing
		response.write("Sorunuz kaydedilmiştir.")
		else
		response.write("Eksik alanlar var. Geriye tıklayıp eksikleri tamamlayınız.")
		end if
%>
<%
	Case "soruduzenle"
Set rs = bag.Execute("SELECT sorular.* FROM sorular where sorular.no="&no&"" ) 
%>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; font-family:Verdana; font-size:10pt; font-weight:bold" bordercolor="#111111" width="750">
<tr>
<td width="100%" valign="middle" align="center">
<form method="POST" style="margin-top: 0; margin-bottom: 0;" action="bilgiyarismasi.asp?bolum=soruguncelle&no=<%=no%>">
<input type="hidden" name="referer" value="<%=Request.ServerVariables("HTTP_REFERER")%>" />
              <label><strong>Tür :
              <input type="text" name="tur" value="<%=rs("grup")%>" />
              </strong></label>
              <strong><br />
              <label>Soru :<br />
              <textarea name="soru" cols="50" rows="4"><%=rs("soru")%></textarea>
              </label>
              <br />
              <label>A :
              <input type="text" name="a" value="<%=rs("a")%>" />
              <input name="dogru" type="radio" value="a" />
              </label>
              <br />
              <label>B :
              <input type="text" name="b" value="<%=rs("b")%>" />
              <input name="dogru" type="radio" value="b" />
              </label>
              <br />
              <label>C :
              <input type="text" name="c" value="<%=rs("c")%>" />
              <input name="dogru" type="radio" value="c" />
              </label>
              <br />
              <label>D :
              <input type="text" name="d" value="<%=rs("d")%>" />
              <input name="dogru" type="radio" value="d" />
              </label><br>Doğru Cevap : <%=rs("dogru")%><br>
              <label>Seviye :
		<select name="seviye">
		<option value="<%=rs("derece")%>"><%=rs("derece")%></option>
		<%
		for b=1 to 25
		%>
		<option value ="<%=b%>"><%=b%></option>
		<%
		next
		%>
		</select>
              </label>
              <br />
		<input type="submit" value="Gönder" />
                </strong>
            </form>
</td>
</tr>
</table>
<%
rs.Close 
Set rs = Nothing
%>
<%
	Case "soruguncelle"
%>
<%
		Function avla(byval hedef) 
		hedef = replace(hedef,"_","") 
		hedef = replace(hedef,"*","")  
		hedef = replace(hedef,"%","") 
		hedef = replace(hedef,"<","") 
		hedef = replace(hedef,">","") 
		hedef = replace(hedef,"chr(13)","<br>") 
		avla=trim(hedef) 
		End Function 
		if len(request.form("tur"))>0 and len(request.form("soru"))>0 and len(request.form("seviye"))>0 and len(request.form("a"))>0 and len(request.form("b"))>0 and len(request.form("c"))>0 and len(request.form("d"))>0 and len(request.form("dogru"))>0 then
		set rs = server.createobject("ADODB.Recordset")
		rs.open "select sorular.* from sorular where sorular.no="&no,bag,1,3
		rs("soru")=avla(request.form("soru"))
		rs("grup")=avla(request.form("tur"))
		rs("a")=avla(request.form("a"))
		rs("b")=avla(request.form("b"))
		rs("c")=avla(request.form("c"))
		rs("d")=avla(request.form("d"))
		rs("dogru")=avla(request.form(request.form("dogru")))
		rs("onay")=true
		rs("derece")=request.form("seviye")
		rs.update
		rs.close
		set rs = nothing
		response.write("Soru güncellenmiştir.<meta http-equiv='Refresh' content='2;url="&request.form("referer")&"'>")
		else
		response.write("Eksik alanlar var. Geriye tıklayıp eksikleri tamamlayınız.")
		end if
%>


<%
	Case "sorusil"
set rs=server.CreateObject("adodb.recordset")
rs.Open "delete from sorular where sorular.no="&no&"",bag,1,3
response.write("Soru silinmiştir.<meta http-equiv='Refresh' content='2;url="&Request.ServerVariables("HTTP_REFERER")&"'>")
%>
<%
	Case "sartir"
set rs=server.CreateObject("adodb.recordset")
rs.Open "Select * from sorular where sorular.no="&no&"",bag,1,3

rs("derece")=request.form("seviye")
rs.update

rs.Close
set rs=nothing
%>
<meta http-equiv="refresh" content="0;url=<%=request.servervariables("HTTP_REFERER")%>">
<%
	Case "soruonayla"
set rs=server.CreateObject("adodb.recordset")
rs.Open "Select * from sorular where sorular.no="&no&"",bag,1,3

rs("onay")=true
rs.update

rs.Close
set rs=nothing
%>
<meta http-equiv="refresh" content="0;url=<%=request.servervariables("HTTP_REFERER")%>">
<%
	Case else
%>
Linkleri Kullanarak İşlemlerinizi Yapınız.
<%
End Select
bag.Close
set bag=nothing
%>
</body>
</html>
<%
else
response.redirect "giris.asp?sayfa=bilgiyarismasi.asp"
end if%>