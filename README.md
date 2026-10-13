# SKF Fatura Veri Portalı

Türkiye'deki 21 elektrik dağıtım bölgesinden gelen fatura çalışma kitaplarını
standartlaştırıp analiz eden, Excel çıktısı üreten Streamlit uygulaması.

## Gereksinimler

- Python 3.14
- `06-Haziran/SKF Başlıkları.xlsx` mapping dosyası (bu depoda bulunur)
- Analiz edilecek aylık EDAŞ çalışma kitapları ve varsa ENO dosyası

Ham fatura dosyaları, müşteri/ETSO bilgileri, ENO dosyaları ve üretilen raporlar
kişisel veri veya ticari bilgi içerebildiğinden Git tarafından izlenmez. Bunları
uygulama klasörüne ya da seçtiğiniz yerel veri klasörüne koyun.

## Kurulum ve başlatma

Depo kökünde PowerShell kullanarak:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r 06-Haziran/requirements.txt
python run_app.py
```

Uygulama varsayılan olarak yalnızca bu bilgisayardan erişilebilen `127.0.0.1`
adresinde açılır. `--port 8510` gibi bir argümanla port seçebilirsiniz:

```powershell
python run_app.py --port 8510
```

Güvenilir yerel ağda paylaşmak için `python run_app.py --lan` kullanılabilir.
`--lan-open` XSRF korumasını kapattığından yalnızca geliştirme/test içindir.

Uygulama ekranında 21 EDAŞ bölge klasörünün bulunduğu ana dizini seçin. Sonuç
Excel dosyasını arayüzdeki indirme düğmesiyle alın.

## Testler

```powershell
python -m pip install -r requirements-dev.txt
Set-Location 06-Haziran
..\.venv\Scripts\python.exe -m pytest -q
```

Birim testleri sentetik girdilerle çalışır. Golden regresyon testi, yerel EDAŞ
çalışma kitapları ve mapping dosyası mevcutsa 21 bölge/2.593 kayıt akışını da
doğrular; bu özel fatura verileri depoya eklenmez.

## Proje yapısı

- `06-Haziran/app.py`: Streamlit arayüzü ve analiz orkestrasyonu
- `06-Haziran/extract_and_compare.py`: fatura dosyalarını okuma, normalize etme,
  birleştirme, doğrulama ve Excel çıktısı üretme
- `06-Haziran/new_mapping_parser.py`: mapping çalışma kitabındaki bölge
  başlıklarını, notları ve düzeltme kurallarını okuma
- `06-Haziran/core/`: sayı/ETSO/bölge normalizasyonu, sütun çözümleme, reaktif
  hesaplama, ENO zenginleştirme ve Excel yazımı
- `06-Haziran/config/`: sabitler ve mapping başlık ekleri
- `06-Haziran/tests/`: birim ve golden regresyon testleri
- `06-Haziran/DOCS/`: tasarım, güvenlik ve mühendislik dokümanları
