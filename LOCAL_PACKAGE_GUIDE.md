# Windows için şirket içi yerel paket

## Kullanıcı deneyimi

Windows paketi kullanıcı hesabı altında çalışır; Python veya Excel kurulumu
istemez. Kullanıcı ZIP'i yerel klasörüne çıkarıp `SKF'yi Başlat.bat` dosyasına
çift tıklar. Başlatıcı penceresi kapanır; uygulama arka planda çalışmaya devam
eder ve varsayılan tarayıcı yerel adresi açar. Dosyalar Windows'un kendi seçme
pencereleriyle alınır. Analiz aynı bilgisayarda yürür.
Kesin Liste seçimi oturum içinde tutulur; uygulama yeniden açıldığında tekrar seçilir.

Uygulamayı kapatmak için `SKF'yi Durdur.bat` dosyasına çift tıklayın. Tarayıcıyı
kapatmak uygulama sürecini durdurmaz. Başlatma sorunlarında günlük dosyası
`%LOCALAPPDATA%\SKF-Yerel\logs` altına yazılır; ayrıntıları görmek için
`SKF'yi Tanılama Modunda Başlat.bat` kullanılabilir.

## Paket üretme

1. Uygulama sürümünü ve referans dosyalarını gözden geçirin.
2. Windows üzerinde Python 3.14 ve proje `.venv` ortamını hazırlayın:

   ```powershell
   py -3.14 -m venv .venv
   .\.venv\Scripts\python.exe -m pip install -r 06-Haziran\requirements.txt
   ```

3. Depo kökünde paketleyiciyi çalıştırın (mevcut `dist` çıktısını koruyarak ayrı
   bir klasöre üretmek için):

   ```powershell
   $env:SKF_PACKAGE_OUTPUT_DIR = "dist\guncelleme"
   .\.venv\Scripts\python.exe tools\build_local_windows_package.py
   ```

4. Oluşan `dist\guncelleme\SKF-Yerel-Paket-1.3.zip` dosyasını şirketin onaylı
   dağıtım kanalına verin. SHA-256 değeri aynı klasördeki `.sha256` dosyasındadır.

Paketleyici yalnızca uygulamanın kaynak kodunu, `SKF Başlıkları.xlsx` dosyasını,
ek kalem bedel aktarım şablonunu, varsayılan Sistem Kullanım Bedelleri kitabını,
gerekli Python bağımlılıklarını ve Python çalışma zamanını alır. Bölge faturaları,
ENO/sayaç listeleri, yerel dosya
yolu tercihleri, çıktı dosyaları, proje testleri ve arşivlenmiş analiz kodları
pakete girmez.

## Şirket içinde dağıtım

- BT ekibi ZIP'i Intune/Software Center veya şirketin onaylı dosya dağıtım alanı
  üzerinden dağıtabilir. Uygulama klasörünü kullanıcı profilinde yazılabilir
  bir konuma çıkarın; örneğin `%LOCALAPPDATA%\SKF-Yerel`.
- Böylece kurulum için yerel yönetici izni gerekmez. Kullanıcılar başlatıcıya
  çift tıklayarak uygulamayı açar. Güncellemede paket klasörü yeni sürümle
  değiştirilir.
- Dağıtım öncesi BT; paketi zararlı yazılım taramasından geçirmeli, kurumun
  kod imzalama/izin listesi sürecini uygulamalı ve gerekiyorsa hash değerini
  onaylamalıdır. İmzalama, kurumun sertifikasıyla BT ortamında yapılmalıdır.
- Uygulama yerel arayüz için `127.0.0.1` adresini kullanır; ortak iş
  istasyonlarında `--lan` seçeneğini etkinleştirmeyin. Başlatıcı Streamlit
  kullanım istatistiklerini kapatır.
- Ham veri dosyalarını uygulama klasörüne veya GitHub deposuna kopyalamayın.
  Kullanıcı onları kendi bilgisayarından seçer; Excel çıktısı tarayıcıdan yerel
  olarak indirilir.

Bu depo kaynak kodu ve paket üreticisini içerir. Şirketin onaylı Windows kod
imzası ve yazılım dağıtım sistemi burada yapılandırılmamıştır.
