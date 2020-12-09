<%
yonetimsifre="admin" 'Yönetici Şifresi Kesinlikle değiştiriniz.
yarismaciadi=request.querystring("yadi") 'Yarışmacı adı değişkeni. Buraya kendi siteminizdeki session değişkenini yazabilirsiniz.
veritabani=server.mappath("/db/denemebilgiyarismasi.mdb") 'Kullanıcı tarafı için Veritabanı yolu
veritabaniyonetim=server.mappath("/db/denemebilgiyarismasi.mdb") 'Yönetici tarafı için Veritabanı yolu
%>