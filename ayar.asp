<%
yonetimsifre="admin" 'Yönetici Þifresi Kesinlikle deðiþtiriniz.
yarismaciadi=request.querystring("yadi") 'Yarýþmacý adý deðiþkeni. Buraya kendi siteminizdeki session deðiþkenini yazabilirsiniz.
veritabani=server.mappath("/db/denemebilgiyarismasi.mdb") 'Kullanýcý tarafý için Veritabaný yolu
veritabaniyonetim=server.mappath("/db/denemebilgiyarismasi.mdb") 'Yönetici tarafý için Veritabaný yolu
%>