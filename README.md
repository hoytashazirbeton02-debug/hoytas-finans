

## v56 güncelleme kaynağı

Güncelleme istemcisi GitHub Releases yanında HTTPS JSON manifest kaynağını da destekler. Manifest alanları: `version`, `package_url`, `sha256`. Manifest ve ZIP adresleri HTTPS olmalı; SHA-256 doğrulanmadan kurulum yapılmaz. Örnek ve ayrıntılar `DEGISIKLIKLER_v56.txt` içindedir.

## v59 notu
v58'deki beton müşteri önerileri ve serbest yazı davranışı korunmuştur. Güncelleme istemcisine manifest/sürüm/HTTPS/SHA-256 kontrolleri güçlendirilmiş; açılışta sessiz kontrolde yeni sürüm kurulmadan önce kullanıcı onayı istenmesi sağlanmıştır. 25 test başarılıdır.

Canlı internet güncellemesi için anonim HTTPS erişimli `manifest.json` ve ZIP yayınlanmalı, ardından `%LOCALAPPDATA%\HOYTASFinans\guncelleme_ayarlari.json` içindeki `manifest_url` bir kez tanımlanmalıdır. Yayın kurulumu `YAYIN_SUNUCUSU_KURULUM_v59.txt` dosyasındadır. Bu paket doğrulanmamış/sahte bir sunucu adresi içermez.