<html>
<head>
<title>Web Kartal� Bilgi Yar��mas�</title>
<script language="JavaScript">
function newWindow(mypage,myname,w,h,features) {
  if(screen.width){
  var winl = (screen.width-w-10);
  var wint = 0;
  }else{winl = 0;wint =0;}
  if (winl < 0) winl = 0;
  if (wint < 0) wint = 0;
  var settings = 'height=' + h + ',';
  settings += 'width=' + w + ',';
  settings += 'top=' + wint + ',';
  settings += 'left=' + winl + ',';
  settings += features;
  win = window.open(mypage,myname,settings);
  win.window.focus();
}
</script>
</head>
<body>
<center>
<input type="text" name="yadi" id="yadi" value="deneme"/><br>
<input type="button" value="Bilgi Yar��mas�" style="font-size: 8pt" onclick="newWindow('bilgiyarismasi.asp?yadi='+document.getElementById('yadi').value,'yarisma',500,450)">
<br><a href="webkartalibilgiyarismasi.zip">Bu Scripti indirmek i�in t�klay�n�z.</a>
</center>
</body>
</html>