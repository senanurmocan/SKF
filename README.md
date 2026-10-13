# CK Enerji SKF Fatura Veri Portalı

Elektrik dağıtım faturalarını birleştiren ve Excel raporu oluşturan Streamlit uygulaması.

## Şirket içi kullanım

Önerilen yayın modeli Windows Server, IIS reverse proxy ve Windows Authentication'dır. Erişim AD güvenlik grubuyla sınırlandırılır; Streamlit yalnızca sunucunun `127.0.0.1:8505` adresinde dinler. Tarayıcı dosya yüklemez. Uygulama, yapılandırılmış şirket paylaşımındaki dönem ve Excel dosyalarını sunucu üzerinden seçtirir.

Dosyalar ağ paylaşımından şirket sunucusuna okunur, sunucu belleğinde işlenir ve sonuç şirket içi HTTPS üzerinden indirilir. Fatura verileri Streamlit Cloud'a veya üçüncü taraf depolamaya gönderilmez. Verilerin sunucu belleğinde işleneceğini ve indirilen raporun kullanıcı cihazına ulaşacağını dikkate alın.

Başlangıç ve BT kurulum adımları için [şirket içi dağıtım kılavuzuna](DEPLOYMENT_INTERNAL.md) bakın. IIS örnek yapılandırması ve servis betikleri `deploy/windows/` altındadır.

## Yerel geliştirme

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r 06-Haziran/requirements.txt
python run_app.py
```

Yerel başlatıcı yalnızca `127.0.0.1` üzerinde çalışır. `--server` şirket içi dağıtım içindir ve yalnızca IIS kimlik doğrulamalı reverse proxy arkasında kullanılmalıdır. `--lan-open` kaldırılmıştır.

## Genel web sayfası

Streamlit Cloud'daki genel sayfa dosya seçmez ve fatura verisi işlemez; yalnızca yerel indirme/yönlendirme ekranıdır.

## Veri ve dosyalar

Fatura kaynakları, müşteri/ETSO bilgileri ve üretilen raporlar Git tarafından izlenmez. Şirket içi sunucudaki uygulama servisine yalnızca gerekli ağ paylaşımında salt okunur yetki verin. Paylaşılan mapping dosyası `06-Haziran/SKF Başlıkları.xlsx` konumundadır.

## Testler

```powershell
python -m pip install -r requirements-dev.txt
Set-Location 06-Haziran
..\.venv\Scripts\python.exe -m pytest -q
```
