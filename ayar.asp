<%
yonetimsifre="admin" 'Y�netici �ifresi Kesinlikle de�i�tiriniz.
yarismaciadi=request.querystring("yadi") 'Yar��mac� ad� de�i�keni. Buraya kendi siteminizdeki session de�i�kenini yazabilirsiniz.
veritabani=server.mappath("/db/denemebilgiyarismasi.mdb") 'Kullan�c� taraf� i�in Veritaban� yolu
veritabaniyonetim=server.mappath("/db/denemebilgiyarismasi.mdb") 'Y�netici taraf� i�in Veritaban� yolu
%>