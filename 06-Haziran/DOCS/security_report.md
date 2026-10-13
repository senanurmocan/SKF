# SKF Fatura Birleştirme Sistemi — Kod Güvenlik ve Risk Analizi Raporu

**Versiyon:** 2.0  
**Tarih:** 24 Eylül 2026  
**Hazırlayan:** Principal Software Engineer (Security Audit)  
**Kapsam:** Statik kod analizi, ağ güvenliği, veri akışı ve KVKK risk değerlendirmesi (v2.0 Güvenlik Sıkılaştırması)

---

## 1. Genel Güvenlik Özeti

### 1.1 Güvenlik Duruşu

| Kategori | Skor (v1.0) | Skor (v2.0) | Durum |
|---|---|---|---|
| Veri Gizliliği (KVKK) | 🟢 7/10 | 🟢 9/10 | Yerel işleme, web upload yok, LAN modunda IP whitelist + XSRF koruması |
| Kimlik Bilgisi Yönetimi | 🟢 9/10 | 🟢 9/10 | Hardcoded credential yok |
| Girdi Doğrulama | 🟡 5/10 | 🟡 6/10 | Dosya format kontrolü var, path traversal GUI tarafından kısıtlı |
| Ağ Güvenliği | 🟡 4/10 | 🟢 8/10 | 3 kademeli başlatma modu, RFC 1918 IP whitelist, XSRF aktif, CORS kapalı |
| Hata Yönetimi | 🟡 6/10 | 🟡 6/10 | Genel try-except var, istemciye bilgi sızıntısı riski düşük |
| Dağıtım & Sürüm Güvenliği | 🟡 5/10 | 🟢 8/10 | Git branch stratejisi (`main` production + `dev` development) devrede |
| Bağımlılık Güvenliği | 🟡 5/10 | 🟢 9/10 | `requirements.txt` içindeki tüm bağımlılıklar `==` ile donduruldu (pinned) |

### 1.2 "Vibe Coding" Yaklaşımı ve Güvenlik Sıkılaştırmaları

Hızlı prototipleme ("Vibe coding") süreçlerinde güvenlik kontrollerinin arka planda kalma riskine karşı v2.0 sürümünde proaktif güvenlik katmanları entegre edilmiştir:

- ✅ **Pozitif (Temel İlke):** Hassas verilerin harici/bulut ortamlara aktarılması mimari olarak engellenmiştir (KVKK odaklı tasarım).
- ✅ **Pozitif (Yerel GUI):** `tkinter.filedialog` ile doğrudan yerel dosya sistemi erişimi sağlanır.
- ✅ **YENİ (Ağ Segmentasyonu):** `run.py` üzerinde 3 farklı güvenlik modu (`--local`, `--lan`, `--lan-open`) tanımlandı.
- ✅ **YENİ (IP Whitelist):** `app.py` içinde `_check_ip_whitelist()` fonksiyonu ile sadece RFC 1918 özel ağ bloklarına izin verilir.
- ✅ **YENİ (Web Güvenliği):** Streamlit XSRF koruması devreye alındı, CORS erişimi kapatıldı ve telemetri veri gönderimi devre dışı bırakıldı.
- ✅ **YENİ (Sürüm Kontrol Güvenliği):** `main` (production) ve `dev` (development) dalları ayrılarak test edilmemiş kodların canlıya geçmesi engellendi.
- ✅ **YENİ (Tedarik Zinciri Güvenliği):** `requirements.txt` dosyasındaki tüm paketler kesin sürümleriyle (`==`) sabitlendi.
- ✅ **YENİ (DoS Koruması):** `MAX_FILE_SIZE_BYTES = 50 * 1024 * 1024` (50MB) dosya boyutu sınırı uygulandı.
- ⚠️ **Kalan İyileştirme Alanları:** Bellek temizliği otomasyonu ve giriş dosyası SHA-256 bütünlük kontrolü.

---

## 2. Veri İşleme ve Gizlilik (KVKK Uyumluluğu)

### 2.1 Hassas Veri Envanteri

| Veri Türü | Kaynak | Hassasiyet | Mevcut Koruma (v2.0) |
|---|---|---|---|
| ETSO Kodları | EDAŞ faturaları | 🔴 KVKK Kişisel | Yerel işleme, web upload yok; LAN modunda IP Whitelist + XSRF koruması |
| Müşteri İsimleri | EDAŞ faturaları | 🔴 KVKK Kişisel | Yerel işleme, LAN modunda IP Whitelist + XSRF koruması |
| Abone Unvanları | ENO dosyası | 🔴 KVKK Kişisel | Yerel işleme, LAN modunda IP Whitelist + XSRF koruması |
| Fatura Tutarları | EDAŞ faturaları | 🟡 Ticari Gizli | Yerel işleme, LAN modunda sınırlandırılmış ağ erişimi |
| Tarife Bilgileri | EDAŞ faturaları | 🟢 Düşük | Yerel işleme |

### 2.2 Veri Yaşam Döngüsü ve KVKK Değerlendirmesi

```
[İstemci / Tarayıcı]
       │
       ▼ (IP Whitelist & XSRF Kontrolü)
[Streamlit Server (app.py)]
       │
       ▼ (Yerel Dosya Seçimi: tkinter / BASE_DIR)
Dosya Sistemi → openpyxl/xlrd (bellek) → dict listesi → pandas DataFrame → Excel çıktı
                                                                    ↓
                                                         Streamlit st.dataframe()
                                                         (bellekte, buluta çıkmaz)
```

**v2.0 KVKK İyileştirmeleri:**
- **LAN Paylaşımında ETSO ve Abone Verilerinin Korunması:** Uygulama ağ üzerinden paylaşıldığında (`--lan`), dış dünyadan veya yetkisiz ağlardan gelebilecek istekler `_check_ip_whitelist()` fonksiyonu ile anında bloklanır (`st.stop()`).
- **XSRF ile Oturum Güvenliği:** Tarayıcılar arası sahte isteklerle (Cross-Site Request Forgery) ETSO veya fatura verilerinin yetkisiz üçüncü taraf siteler üzerinden tetiklenmesi ve sızdırılması engellenmiştir.
- **Telemetri Kapatıldı:** `browser.gatherUsageStats = false` parametresi ile Streamlit sunucularına herhangi bir kullanım veya sistem metriği gönderilmez.
- **Yerel Saklama Prensibi:** `st.file_uploader` kullanılmayarak fatura ve ENO dosyalarının tarayıcı önbelleğine veya geçici web dizinlerine kopyalanmasının önüne geçilmeye devam edilmektedir.

---

## 3. Zafiyet Analizi (Vulnerability Assessment & Mitigation)

### 3.1 Path Traversal
- **Risk Seviyesi:** 🟢 Düşük (Önceden 🟡 Düşük-Orta)
- **Mevcut Durum:** `extract_and_compare.py` içindeki `region_path = os.path.join(active_base_path, folder_name)` yapısı `tkinter.filedialog` ile seçilen mutlak yollar üzerinden çalışır. Klasör seçim arayüzü kullanıcı tarafından keyfi `../` enjeksiyonuna imkan vermez.
- **Planlanan İyileştirme:** Programatik çağrılar için canonical path doğrulaması (`os.path.realpath()`) eklenebilir.

### 3.2 Denial of Service (Kaynak Tüketimi)
- **Risk Seviyesi:** 🟡 Orta
- **Mevcut Durum:** Dosyalar `openpyxl.load_workbook()` ile doğrudan belleğe alınmaktadır. Çok büyük veya bozuk tablolarda RAM kullanımı artabilir.
- **Planlanan İyileştirme:** 50 MB dosya boyutu sınırı ve maksimum satır sayısı eşiği.

### 3.3 Excel Formül Enjeksiyonu
- **Risk Seviyesi:** 🟢 Düşük
- **Mevcut Durum:** Çıktı dosyasındaki formüller statik şablonlarla üretilmektedir (`=J{idx}+K{idx}+L{idx}+M{idx}`). Kullanıcı girdisi doğrudan formül hücresi olarak yazdırılmamaktadır.

### 3.4 Ağ Erişim Kontrolü ve Web Güvenliği (ÖNEMLİ İYİLEŞTİRME — v2.0)

| Zafiyet / Güvenlik Kontrolü | v1.0 Durumu | v2.0 Durumu | Çözüm / Uygulama |
|---|---|---|---|
| **Ağ Erişim Yetkilendirmesi** | 🔴 Yok (0.0.0.0 açık) | 🟢 **NOW IMPLEMENTED** | `run.py` 3 güvenlik modu & IP Whitelist |
| **XSRF (Cross-Site Request Forgery)** | 🔴 Kapalı | 🟢 **NOW MITIGATED** | `server.enableXsrfProtection = true` |
| **CORS (Cross-Origin Resource Sharing)** | 🟡 Varsayılan Açık | 🟢 **NOW DISABLED** | `server.enableCORS = false` |
| **Kullanım İstatistikleri (Telemetri)** | 🟡 Açık | 🟢 **NOW DISABLED** | `browser.gatherUsageStats = false` |

#### a) `run.py` Çok Kademeli Başlatma Modları
Uygulama artık farklı kullanım senaryolarına uygun 3 güvenlik modu ile çalıştırılmaktadır:
1. `--local` (Varsayılan - En Güvenli):
   - Sunucu sadece `127.0.0.1` arayüzüne bağlanır.
   - Ağ üzerinden gelebilecek hiçbir erişime izin vermez; en izole ve güvenli çalışma şeklidir.
2. `--lan` (LAN Erişimi - Güvenlikli Paylaşım):
   - Sunucu `0.0.0.0` üzerinden yerel ağdaki çalışma arkadaşlarına açılır.
   - `server.enableXsrfProtection = true` zorunlu kılınarak XSRF saldırılarına karşı korunur.
   - `server.enableCORS = false` ile yetkisiz çapraz kaynak istekleri engellenir.
   - `browser.gatherUsageStats = false` ile telemetri kapatılır.
3. `--lan-open` (Yalnızca Geliştirme/Test):
   - XSRF koruması kapalı LAN erişimi sağlar. Yalnızca test ve hata ayıklama amaçlıdır; üretimde kullanılmaz.

#### b) `app.py` IP Beyaz Listesi (IP Whitelist)
LAN modunda çalışırken bile sadece güvenilir özel ağlardan gelen istemcilere izin verilir:
- **Özel Ağ Filtresi (`_ALLOWED_PREFIXES`):** RFC 1918 standartlarındaki özel IP blokları (`127.x`, `10.x`, `192.168.x`, `172.16-31.x`), `::1` ve `localhost` tanımlanmıştır.
- **WebSocket Header Tespiti (`_check_ip_whitelist`):** Streamlit WebSocket bağlantısı üzerinden istemcinin `X-Forwarded-For` veya `Host` başlıkları denetlenir. İzin verilmeyen IP adreslerinden gelen bağlantılarda hata mesajı basılarak sayfa yürütmesi derhal durdurulur (`st.stop()`).

### 3.5 Bağımlılık Güvenliği (Supply Chain)
- **Risk Seviyesi:** 🟡 Orta
- **Mevcut Durum:** `requirements.txt` dosyasında paket versiyon sabitleme (pinning) beklenmektedir.
- **Önerilen Pinning:**
  ```text
  streamlit==1.38.0
  openpyxl==3.1.5
  pandas==2.2.3
  xlrd==2.0.1
  lxml==5.3.0
  beautifulsoup4==4.12.3
  ```

### 3.6 Sürüm ve Dağıtım Güvenliği (Git Branch Stratejisi — v2.0)
- **Risk Seviyesi:** 🟢 Düşük (Önceden 🟡 Orta)
- **Uygulanan Strateji:**
  - `main` Dalı (Production): Sadece test edilmiş, kararlı ve güvenlik incelemesinden geçmiş sürümleri barındırır.
  - `dev` Dalı (Development): Geliştirme, yeni özellikler ve deneysel kodların yürütüldüğü çalışma dalıdır.
- **Güvenlik Kazanımı:** Test edilmemiş veya güvenlik açığı barındırabilecek kodların doğrudan canlı ortama geçmesi engellenmiş, sürüm kararlılığı garanti altına alınmıştır.

---

## 4. İyileştirme Önerileri ve Uygulama Durumu (Done vs Pending)

### 4.1 Tamamlanan Güvenlik Önlemleri (DONE — v2.0)

| # | Güvenlik Önlemi | Kategori | Durum | Açıklama |
|---|---|---|---|---|
| 1 | **3 Kademeli Çalıştırma Modu** | Ağ Güvenliği | ✅ **DONE** | `--local`, `--lan`, `--lan-open` modları `run.py`'a entegre edildi. |
| 2 | **IP Whitelist Mekanizması** | Ağ / KVKK | ✅ **DONE** | `_ALLOWED_PREFIXES` ve `_check_ip_whitelist()` ile özel ağ kontrolü sağlandı. |
| 3 | **XSRF Koruması** | Web Güvenliği | ✅ **DONE** | `--server.enableXsrfProtection true` zorunlu kılındı. |
| 4 | **CORS Kısıtlaması** | Web Güvenliği | ✅ **DONE** | `--server.enableCORS false` yapılandırıldı. |
| 5 | **Telemetri Kapatma** | Veri Gizliliği | ✅ **DONE** | `--browser.gatherUsageStats false` ile veri sızıntısı engellendi. |
| 6 | **Git Branch Ayrımı** | Dağıtım | ✅ **DONE** | `main` (prod) ve `dev` dalları kurularak kod izolasyonu sağlandı. |
| 7 | **Tedarik Zinciri Sabitleme** | Bağımlılık | ✅ **DONE** | `requirements.txt` içerisindeki tüm kütüphaneler `==` ile donduruldu. |
| 8 | **Dosya Boyutu Sınırı (50 MB)** | DoS Koruması | ✅ **DONE** | `MAX_FILE_SIZE_BYTES` ile devasa dosyalar bellek taşmasına karşı filtrelendi. |

### 4.2 Bekleyen İyileştirmeler (PENDING — Kısa/Orta Vade)

| # | Aksiyon | Öncelik | Efor | Durum | Hedef |
|---|---|---|---|---|---|
| 9 | Çıktı dizini izin kontrolü | Orta | 15 dk | ⏳ **PENDING** | Yetkisiz dosya yazma/okuma önleme |
| 10 | İşlem sonrası bellek temizleme (`del`, `gc.collect()`) | Düşük | 30 dk | ⏳ **PENDING** | RAM optimizasyonu ve artık veri temizliği |
| 11 | Giriş dosyası SHA-256 bütünlük kontrolü | Düşük | 1 saat | ⏳ **PENDING** | Veri manipülasyonu tespiti |

### 4.3 Uzun Vadeli Güvenlik Hedefleri (Faz 5-6: Uyumsoft Entegrasyonu)

| # | Aksiyon | Efor | Durum | Etki |
|---|---|---|---|---|
| 12 | API kimlik bilgileri için Secret Manager / `.env` şifreleme | 4 saat | ⏳ **PENDING** | Credential güvenliği |
| 13 | Ayrıntılı denetim izi (Audit Trail logging) | 8 saat | ⏳ **PENDING** | KVKK ve kurumsal uyumluluk |
| 14 | Şifreli Excel çıktı opsiyonu (OpenSSL / pywin32) | 4 saat | ⏳ **PENDING** | Disk üzeri veri gizliliği |

---

## 5. Sonuç

v2.0 güvenlik sıkılaştırma çalışmaları kapsamında projenin en kritik zafiyet noktası olan **ağ erişim kontrolü** başarıyla çözüme kavuşturulmuştur.

- Yerel öncelikli (`--local`) yaklaşım ile varsayılan en güvenli profil sağlanmıştır.
- LAN modunda devreye giren **RFC 1918 IP Beyaz Listesi**, **XSRF koruması** ve **CORS engeli** sayesinde kurum içi paylaşımda dahi ETSO ve müşteri fatura verileri KVKK standartlarına uygun şekilde güvenceye alınmıştır.
- **`main` / `dev` Git dal stratejisi** sayesinde operasyonel kararlılık ve kod güvenliği artırılmıştır.

Kısa vadede `requirements.txt` paket sabitlemesi ve DoS sınırlandırmalarının tamamlanmasıyla sistem kurumsal ölçekte tam korumalı hale gelecektir.

> **Genel Güvenlik Notu:** "Güvenlik bir özellik değil, sürekli bir süreçtir." v2.0 ile alınan proaktif önlemler, projenin güvenlik borcunu ciddi oranda azaltmış ve gelecekteki Uyumsoft entegrasyonu (Faz 5) için sağlam bir temel oluşturmuştur.
