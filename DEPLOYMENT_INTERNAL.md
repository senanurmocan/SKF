# Şirket içi sunucu dağıtımı

Bu dağıtım, uygulamayı Windows Server üzerinde çalıştırır. Kullanıcılar şirket içi HTTPS adresine gider; IIS Windows Authentication ile etki alanı hesabını doğrular ve yalnızca yetkili Active Directory grubuna erişim verir. Streamlit uygulaması yalnızca sunucunun `127.0.0.1:8505` adresinde dinler.

## Veri akışı ve erişim sınırı

- Tarayıcıdan Excel dosyası yükleme yoktur. Uygulama, sunucu servis hesabının okuyabildiği sabit ağ paylaşımı klasörlerinden dosya seçtirir.
- Dosyalar ağ paylaşımından uygulama sunucusuna okunur ve sunucu belleğinde işlenir. İşlenen tablo/rapor oturum belleğinde tutulur; sonuç indirilirken şirket içi HTTPS üzerinden kullanıcının tarayıcısına gider.
- Uygulama fatura dosyalarını genel web servisine veya üçüncü taraf depolamaya göndermez. Bu tasarım “hiçbir yere kopyalanmama” anlamına gelmez: sunucu dosyayı işler ve sonuç, indiren kullanıcının cihazında bulunur.
- Uygulama servisine yalnızca gereken paylaşımlarda salt okunur yetki verin. Uygulama erişim grubunu küçük tutun; grup üyeleri yapılandırılmış kaynaklardaki dosyaları kullanabilir.
- Raporlar sunucuda kalıcı dosya olarak yazılmaz. Kullanıcı indirdiğinde rapor kendi tarayıcısına iner. Paylaşımdaki girdiler değiştirilmez.
- Oturumu kapatmak için tarayıcı sekmesini kapatmak yeterli olmayabilir; Streamlit kopan oturumların temizliğini 120 saniyelik TTL sonrasında yapabilir. Servisin yeniden başlatılması bütün oturum belleğini temizler.

## Windows Server hazırlığı

1. Uygulamayı yalnızca şirketin yönettiği, domain üyesi bir Windows Server'a kurun. Bir iç DNS adı ve şirket içi CA tarafından verilen TLS sertifikası kullanın.
2. IIS'e **Windows Authentication**, **URL Authorization**, **WebSocket Protocol**, **URL Rewrite 2** ve **Application Request Routing (ARR)** bileşenlerini kurun. ARR proxy özelliğini etkinleştirin ve dışarıdan gelen `Host` başlığını koruyun. IIS reverse proxy ve WebSocket ayarları Streamlit'in kullandığı HTTPS adresine göre yapılmalıdır.
3. `deploy/windows/web.config` dosyasını IIS sitesine kopyalayın. `CKENERJI\SKF-Portal-Users` değerini gerçek, erişimi onaylanmış AD güvenlik grubuyla değiştirin. Tek tek kullanıcı adları uygulamaya yazılmaz. IIS'te Anonymous Authentication kapalı, Windows Authentication açık olmalıdır. İzin kuralı bütün HTTP ve WebSocket yollarına uygulanmalıdır.
4. IIS sitesine geçerli iç TLS sertifikası bağlayın. Güvenlik duvarında kullanıcılar için yalnızca gerekli HTTPS erişimini açın; Streamlit'in `8505` portuna ağdan erişim izni vermeyin. IIS ve Streamlit aynı sunucuda çalışmalıdır.
5. Proje sanal ortamını ve sabitlenmiş `06-Haziran/requirements.txt` bağımlılıklarını şirketin onaylı yazılım kaynağından kurun. Kurulumu `deploy/windows/Install-SKF-InternalServer.ps1` ile yapabilir veya onaylı wheelhouse klasörünü `-Wheelhouse` parametresiyle verebilirsiniz. Çalışma anında genel internete erişim gerekmez. Örnek: `powershell -NoProfile -ExecutionPolicy Bypass -File .\deploy\windows\Install-SKF-InternalServer.ps1 -Wheelhouse C:\ApprovedWheelhouse`.
6. `deploy/windows/server.env.example` içindeki örnekleri sunucudaki gerçek paylaşım yollarıyla değiştirin. Bu değerleri kaynak depoya değil Windows servisinin ortamına girin. `SKF_DATA_ROOT`, seçilecek ay klasörlerini doğrudan içermelidir. Temmuz 2026 için kullanıcı örneğindeki paylaşım yolunun `...\2026` ile biten üst klasörünü verin; içindeki `7-Temmuz` seçilebilir. Diğer `*_ROOT` ayarları doğrudan listelenecek Excel dosyalarını içermelidir.
7. Streamlit sürecini `deploy/windows/Start-SKF-InternalServer.ps1` ile BT'nin onayladığı bir Windows servis yöneticisi altında çalıştırın. Servis için ayrı bir yönetilen domain servis hesabı (tercihen gMSA) kullanın; bu hesaba uygulama dosyalarında çalıştırma/okuma ve yalnızca gerekli fatura paylaşımlarında salt okunur yetki verin. Kullanıcı hesabı parolalarını script'e veya `.env` dosyasına yazmayın.
8. `SKF_PUBLIC_URL` ve `SKF_ALLOWED_HOSTS` değerlerinin IIS'teki HTTPS adıyla aynı olduğunu doğrulayın. Uygulama bu adı Streamlit WebSocket host kontrolünde kullanır.

Kullanıcının domain hesabı IIS tarafından doğrulanır; uygulama, ekrana yazılan Windows kullanıcı adını kimlik bilgisi veya parola olarak kullanmaz. AD grubunun üyeliğini BT yönetir. Sunucudaki SMB erişimi ise kullanıcıların kişisel hesaplarıyla değil, kısıtlı uygulama servis hesabıyla yapılır.

`SKF_OUR_INVOICES_ROOT` isteğe bağlıdır. `SKF_ENO_ROOT` verilmezse Kesin Liste, kullanıcının seçtiği dönem klasörünün içinde otomatik aranır. Uygulama seçim kutularına tam sunucu yolunu göndermez; seçim değerleri sunucu tarafındaki indekslerle eşleştirilir. Her iki listede de yalnızca yapılandırılmış klasörün doğrudan içindeki Excel dosyaları görünür.

## Yayına almadan önce

- Yetkisiz bir domain kullanıcısının IIS'ten `401/403` aldığını ve izinli AD grubundaki bir kullanıcının uygulamayı açabildiğini doğrulayın.
- Sunucu dışından `8505` portuna bağlantının başarısız olduğunu doğrulayın.
- IIS log erişimini yalnızca yöneticilere verin ve kurumun kayıt saklama süresini uygulayın. Uygulama loglarında müşteri verisi tutulması gerekmez.
- Servis hesabının örnek dönem klasörünü okuyabildiğini, yazamadığını doğrulayın.
- Raporun indirilmesinin çalıştığını ve istemcinin HTTPS kullandığını doğrulayın.

Bu depo, IIS yapılandırma örneğini ve uygulama servis başlatıcısını sağlar. Gerçek IIS sitesi, TLS sertifikası, AD grubu, Windows servis hesabı ve dosya paylaşımı izinleri şirket sunucusunda BT/yetkili yönetici tarafından tanımlanmalıdır.
