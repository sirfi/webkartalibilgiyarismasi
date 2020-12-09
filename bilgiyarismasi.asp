<%@LANGUAGE="VBSCRIPT" CODEPAGE="1254"%>
<!--#include file ="ayar.asp"-->
<%
if len(Session("byadi"))=0 then
Session("byadi")=yarismaciadi
end if
if len(Session("byadi"))>1 then
Set  bag = server.createobject("ADODB.Connection" )  
bag.open("DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & veritabani)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-9" />
<title>Bilgi Yarışması</title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style4 {
	font-size: 48px;
	font-family: Verdana, Arial, Helvetica, sans-serif;
}
.link {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 14px;
	text-decoration: none;
	color: #000000;
}
.style10 {font-size: 16px; font-weight: bold; }
.style13 {font-size: 12px}
-->
</style>
</head>

<body>
<table width="500" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="400"><img src="resimler/bilgiyarisma.PNG" width="400" height="75" /></td>
    <td width="100" align="center" valign="middle"><span class="style4" id="sayac">30</span>
<%if request.querystring("bolum")="sorucevapla" then%>
<script language="javascript">
<!--
var zaman=31
function gerisayim(){
if (zaman!=0){
zaman-=1
document.getElementById('sayac').innerHTML=zaman
}
else{
document.getElementById('soruform').submit()
return
}
setTimeout("gerisayim()",1000)
}
gerisayim()
//-->
</script>
<%end if%>
	</td>
  </tr>
  <tr>
    <td height="350" colspan="2" align="center" valign="middle">
	<%
	select case request.querystring("bolum")
		case "","anasayfa"
	session("seviye")=1
	session("soru")=0
	session("puan")=0
	session("dsayisi")=0
	session("kumbara")=0
	session("atla")=0
	session("degistir")=0
	session("cekil")=0
	%>
	<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" id="anasayfa">
      <tr>
        <td width="15" height="15">&nbsp;</td>
        <td height="15">&nbsp;</td>
        <td width="15" height="15">&nbsp;</td>
      </tr>
      <tr>
        <td width="15">&nbsp;</td>
<%
randomize 
rsayi=int((rnd * 10000)+ 0)
%>
        <td align="center" valign="middle"><%if len(request.querystring("bolum"))=0 then%>Web Kartalı Bilgi Yarışmasına Hoşgeldiniz<br />Sayın <%=Session("byadi")%><br /><%end if%>
          <a href="?bolum=sorucevapla&rsayi=<%=rsayi%>" class="link">Başlamak İçin Tıklayınız.</a><br />
          <br />
          <a href="?bolum=sorucevapla&rsayi=<%=rsayi%>" target="_self"><img src="resimler/yarismabasla.gif" alt="Yarışmaya başlamak için tıklayınız." border="0" /></a> </td>
        <td width="15">&nbsp;</td>
      </tr>
      <tr>
        <td width="15" height="15">&nbsp;</td>
        <td height="15">&nbsp;</td>
        <td width="15" height="15">&nbsp;</td>
      </tr>
    </table>
	<%
		case "sorucevapla"
	if session("soru")=session("seviye") then
	%>
	<meta http-equiv="refresh" content="0;URL=?bolum=anasayfa">
	<%
	else
	session("soru")=session("seviye")
	set rs = server.createobject("ADODB.Recordset")
	rs.open "select sorular.* from sorular where onay=true and derece="&session("seviye"),bag,1,3
	if not rs.eof then
	randomize 
	rs.move(int((rs.recordcount * rnd)+ 0)) 
	%>
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" id="sorucevapla">
      <tr>
        <td width="15" height="15">&nbsp;</td>
        <td width="40">&nbsp;</td>
        <td width="15">&nbsp;</td>
        <td width="370" height="15">&nbsp;</td>
        <td width="15">&nbsp;</td>
        <td width="30" height="15">&nbsp;</td>
        <td width="15" height="15">&nbsp;</td>
      </tr>
      <tr>
        <td width="15">&nbsp;</td>
        <td width="40" align="center" valign="middle">
		<%if session("atla")=0 then%>
		<a href="#no" onclick="document.getElementById('joker').value='atla';document.getElementById('soruform').submit()">
		<img src="resimler/atla.PNG" alt="Atla:Soru doğru bilinmiş sayılır ve sonraki soruya geçilir." width="38" height="38"/></a><%end if%><%if session("degistir")=0 then%><a href="#no" onclick="document.getElementById('joker').value='degistir';document.getElementById('soruform').submit()">
		<img src="resimler/degistir.PNG" alt="Değiştir:Aynı seviyeden başka soru sorulur." width="38" height="38"/></a><%end if%><%if session("cekil")=0 then%><a href="#no" onclick="document.getElementById('joker').value='cekil';document.getElementById('soruform').submit()">
		<img src="resimler/cekil.PNG" alt="Çekil:Puan kaybı olmadan yarışayı bitirilir." width="38" height="38"/></a><%end if%></td>
        <td width="15">&nbsp;</td>
        <td width="370" align="center" valign="middle">
<%
randomize 
rsayi=int((rnd * 10000)+ 0)
%>
	<form action="?bolum=kontrol&rsayi=<%=rsayi%>" method="post" name="soru" id="soruform" style="height:100%">
		<input name="joker" type="hidden" id="joker" value="" />
		<input name="soruno" type="hidden" id="soruno" value="<%=rs("no")%>" />
          -<%=rs("grup")%>-Gönderen : <%=rs("ekleyen")%><br /><%=rs("soru")%><br />
            <label>
              <input type="radio" name="cevap" value="<%=rs("a")%>" onclick="document.getElementById('cevapla').disabled=false"/>
              <%=rs("a")%></label>
            <br />
            <label>
              <input type="radio" name="cevap" value="<%=rs("b")%>" onclick="document.getElementById('cevapla').disabled=false"/>
              <%=rs("b")%></label>
            <br />
            <label>
              <input type="radio" name="cevap" value="<%=rs("c")%>" onclick="document.getElementById('cevapla').disabled=false"/>
              <%=rs("c")%></label>
            <br />
            <label>
              <input type="radio" name="cevap" value="<%=rs("d")%>" onclick="document.getElementById('cevapla').disabled=false"/>
              <%=rs("d")%></label><br />
            <div align="right">
                <input id="cevapla" name="" type="submit" value="Cevapla" disabled/>
&nbsp;&nbsp;&nbsp;&nbsp;            </div>
        </form></td>
        <td width="15">&nbsp;</td>
        <td width="30" align="center" valign="middle">
		<%
		for a=9+session("seviye") to session("seviye") step -1
		if session("seviye")=a then
		renk="#f62817"
		end if
		if session("seviye")<a then
		renk="#82807e"
		end if
		if (a mod 5)<>0 then
		response.Write("<font color='"&renk&"'>"&a&"</font><br>")
		else
		response.Write("<font color='"&renk&"'>-"&a&"-</font><br>")
		end if
		next
		%>
		</td>
        <td width="15">&nbsp;</td>
      </tr>
      
      <tr>
        <td width="15" height="15">&nbsp;</td>
        <td width="40">&nbsp;</td>
        <td width="15">&nbsp;</td>
        <td width="370" height="15">&nbsp;</td>
        <td width="15">&nbsp;</td>
        <td width="30" height="15">&nbsp;</td>
        <td width="15" height="15">&nbsp;</td>
      </tr>
    </table>
	<%
	end if
	rs.close
	set rs = nothing
	end if
		case "kontrol"
	if session("soru")=session("seviye") then
	atla=0
	degistir=0
	cekil=0
	if len(request.form("joker"))>0 then
		select case request.form("joker")
			case "atla"
			if session("atla")=0 then
			session("atla")=1
			atla=1
			end if
			case "degistir"
			if session("degistir")=0 then
			session("degistir")=1
			degistir=1
			end if
			case "cekil"
			if session("cekil")=0 then
			session("cekil")=1
			cekil=1
			end if
			case else
		end select
	end if
	if degistir=1 then
	session("soru")=session("soru")-1
	%>
	<meta http-equiv="refresh" content="0;URL=?bolum=sorucevapla">
	<%
	else
	set srs = server.createobject("ADODB.Recordset")
	srs.open "select sorular.* from sorular where sorular.no="&request.Form("soruno"),bag,1,3
	srs("csayisi")=srs("csayisi")+1
	srs.update
	if srs("dogru")=request.Form("cevap") or atla=1 then
	srs("bsayisi")=srs("bsayisi")+1
	srs.update
	cevap=1
	session("kumbara")=session("kumbara")+session("seviye")
	if (session("seviye") mod 5)=0 then
	session("puan")=session("puan")+session("kumbara")
	session("kumbara")=0
	end if
	session("dsayisi")=session("dsayisi")+1
	else
	cevap=0
	end if
	srs.close
	set srs = nothing
	if session("seviye")=25 or cevap=0 or cekil=1 then
	set yrs = server.createobject("ADODB.Recordset")
	yrs.open "select yarismacilar.* from yarismacilar where yarismacilar.yadi='"&session("byadi")&"'",bag,1,3
	if yrs.eof then
	yrs.addnew
	yrs("yadi")=Session("byadi")
	yrs("puan")=session("puan")
	if cekil=1 then
	yrs("puan")=yrs("puan")+session("kumbara")
	end if
	yrs("sontarih")=""&date()&""
	yrs("hak")=1
	yrs("csayisi")=session("seviye")
	yrs("dsayisi")=session("dsayisi")
	yrs.update
	else
	yrs("puan")=yrs("puan")+session("puan")
	if cekil=1 then
	yrs("puan")=yrs("puan")+session("kumbara")
	end if
	if yrs("sontarih")=""&date()&"" then
	yrs("hak")=yrs("hak")+1
	else
	yrs("hak")=1
	end if
	yrs("sontarih")=""&date()&""
	yrs("csayisi")=yrs("csayisi")+session("seviye")
	yrs("dsayisi")=yrs("dsayisi")+session("dsayisi")
	yrs.update
	end if
	yrs.close
	set yrs = nothing
	end if
	if cevap=1 and not session("seviye")=25 then
	session("seviye")=session("seviye")+1
	%>
<%
randomize 
rsayi=int((rnd * 10000)+ 0)
%>
	Cevap Doğru
	<meta http-equiv="refresh" content="3;URL=?bolum=sorucevapla&rsayi=<%=rsayi%>">
	<%
	end if
	if session("seviye")=25 or cevap=0 or cekil=1 then
	if session("seviye")=25 then
	%>
	Yarışma Bitti
	<%
	end if
	if cevap=0 and cekil=0 then
	%>
	Cevap Yanlış
	<%
	end if
	%>
	<meta http-equiv="refresh" content="3;URL=?bolum=anasayfa">
	<%
	end if
	end if
	end if
		case "durum"
	%>
	  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" id="durum">
        <tr>
          <td width="15" height="15">&nbsp;</td>
          <td height="15">&nbsp;</td>
          <td width="15" height="15">&nbsp;</td>
        </tr>
        <tr>
          <td width="15">&nbsp;</td>
          <td align="center" valign="middle">
		  <%
		  response.write Session("byadi")&"<br />"
		  set yrs = bag.execute("select yarismacilar.* from yarismacilar where yadi='"&Session("byadi")&"'")
		  if not yrs.eof then
		  response.write "Puan:"&yrs("puan")&"<br>"
		  response.write "Bugün "&yrs("hak")&" kere yarıştınız.<br>"
		  response.write "Yarışmaya katıldığınızdan beri "&yrs("csayisi")&" sorudan "&yrs("dsayisi")&" tanesini <br>doğru olarak cevapladınız."
		  end if
		  yrs.close
		  set yrs = nothing
		  %>
		  </td>
          <td width="15">&nbsp;</td>
        </tr>
        <tr>
          <td width="15" height="15">&nbsp;</td>
          <td height="15">&nbsp;</td>
          <td width="15" height="15">&nbsp;</td>
        </tr>
      </table>
	<%
		case "ilk10"
	%>
	  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" id="ilk10">
        <tr>
          <td width="15" height="15">&nbsp;</td>
          <td height="15">&nbsp;</td>
          <td width="15" height="15">&nbsp;</td>
        </tr>
        <tr>
          <td width="15">&nbsp;</td>
          <td align="center" valign="middle"><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="8%" align="center" valign="middle"><span class="style10">Sıra</span></td>
              <td width="39%" align="center" valign="middle"><span class="style10">Adı</span></td>
              <td width="19%" align="center" valign="middle"><span class="style10">Puan</span></td>
              <td width="17%" align="center" valign="middle"><span class="style10">C. Sayisi </span></td>
              <td width="17%" align="center" valign="middle"><span class="style10">D. Sayisi </span></td>
            </tr>
			<%
			set rs = bag.execute("select top 10 yarismacilar.* from yarismacilar order by yarismacilar.puan desc")
			if not rs.eof then
			for x=1 to 10
			%>
            <tr>
              <td align="center" valign="middle"><%=x%></td>
              <td align="center" valign="middle"><%=rs("yadi")%></td>
              <td align="center" valign="middle"><%=rs("puan")%></td>
              <td align="center" valign="middle"><%=rs("csayisi")%></td>
              <td align="center" valign="middle"><%=rs("dsayisi")%></td>
            </tr>
			<%
			rs.movenext
			if rs.eof then
			exit for
			end if
			next
			end if
			rs.close
			set rs = nothing
			%>
          </table></td>
          <td width="15">&nbsp;</td>
        </tr>
        <tr>
          <td width="15" height="15">&nbsp;</td>
          <td height="15">&nbsp;</td>
          <td width="15" height="15">&nbsp;</td>
        </tr>
      </table>
	<%
			case "sorugonder"
	%>
	  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" id="sorugonder">
        <tr>
          <td width="15" height="15">&nbsp;</td>
          <td height="15">&nbsp;</td>
          <td width="15" height="15">&nbsp;</td>
        </tr>
        <tr>
          <td width="15">&nbsp;</td>
          <td align="center" valign="middle">
            <form id="sorugonder" name="sorugonder" method="post" action="?bolum=sorukaydet">
			<span class="style13">Formu doldururken eksik bırakmayınız. Çünkü eksik olan sorulara onay verilmez. <br />Doğru cevabı belirtmek için yanındaki tuşa tıklayınız. </span><br />
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
          <td width="15">&nbsp;</td>
        </tr>
        <tr>
          <td width="15" height="15">&nbsp;</td>
          <td height="15">&nbsp;</td>
          <td width="15" height="15">&nbsp;</td>
        </tr>
      </table>
	<%
		case "sorukaydet"
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
		rs("ekleyen")=Session("byadi")
		rs("onay")=false
		rs("bsayisi")=0
		rs("derece")=request.form("seviye")
		rs("csayisi")=0
		rs.update
		rs.close
		set rs = nothing
		response.write("Sorunuz kaydedilmiştir. Onaylandıktan sonra yayınlanacaktır.")
		else
		response.write("Eksik alanlar var. Geriye tıklayıp eksikleri tamamlayınız.")
		end if
		case else
	end select
	%>
    </td>
  </tr>
  <tr>
    <td colspan="2"><table width="100%" height="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000">
      <tr>
        <td width="17%" align="center" valign="middle"><a href="?bolum=anasayfa" class="link">Anasayfa</a></td>
        <td width="26%" align="center" valign="middle"><a href="?bolum=sorugonder" class="link">Soru Gönder</a> </td>
        <td width="25%" align="center" valign="middle"><a href="?bolum=durum" class="link">Durumunuz</a></td>
        <td width="16%" align="center" valign="middle"><a href="?bolum=ilk10" class="link">İlk 10 </a></td>
        <td width="16%" align="center" valign="middle"><a href="javascript:window.close()" class="link">Çıkış</a></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
<%
bag.close
set bag = nothing
end if
%>