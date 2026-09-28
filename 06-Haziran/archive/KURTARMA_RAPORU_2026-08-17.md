# SKF Proje Kurtarma ve Bütünlük Raporu

Rapor zamanı: 17 Ağustos 2026 11:02 (Europe/Istanbul)

## Sonuç

Üretim uygulaması kurtarıldı ve yeniden çalışır duruma getirildi. Streamlit arayüzü tekrar doğrulanmış `extract_and_compare.py` çekirdeğini kullanıyor. Gerçek 21 bölgelik veri setinde son sağlam sürümle aynı sonuç elde edildi:

- 2.593 nihai kayıt
- 21/21 dağıtım bölgesi
- 353.348.178,79 TL toplam dağıtım bedeli
- Vangölü EDAŞ: 9 kayıt
- Eksik bölge: yok
- Dosya işleme hatası: yok
- Canonical veri SHA-256: `239a50d6ab0f5d9152d9b1fced7a9536788d1ccf1a384760dbe00bef5d21955e`

## Ne Oldu?

Eski `<REPO_ROOT>` yolu artık bulunmuyor; proje aynı oluşturulma zamanı korunarak `<REPO_ROOT>` konumuna taşınmış veya yeniden adlandırılmış görünüyor.

12 Ağustos 2026 saat 16:16:26'da dokümantasyon ve çok sayıda Python yardımcı dosyası aynı saniyede 0 bayta inmiş. Bu, tek tek kullanıcı düzenlemesinden çok toplu/otomatik bir overwrite davranışına işaret ediyor. Saat 16:17 civarında `extract_and_compare.py`, çalışan büyük çekirdek yerine 409 baytlık `core.engine` uyumluluk sarmalayıcısına çevrilmiş; `app.py` de aynı modüler çekirdeğe bağlanmış. Bu değişiklik uygulamada `return_stats` parametre hatasına ve veri paritesi kaybına yol açmış.

Saat 16:19'da bozuk durum `ede7f45` başlangıç commit'i olarak Git'e alınmış ve uzak depoya gönderilmiş. Bu nedenle Git geçmişinde hasar öncesi sürüm bulunmuyor.

İncelenen Codex oturumlarında, PowerShell geçmişinde, Defender kayıtlarında ve GitHub Desktop günlüklerinde dosyaları 0 bayta indiren kesin komut veya süreç bulunamadı. Son sağlam Codex işlemi ile toplu sıfırlama arasında yaklaşık 30 dakika var. Bu nedenle fail kesin olarak bir uygulamaya veya kişiye atanamaz. GitHub Desktop bozuk durumu commit/push etmiş, ancak dosyaları sıfırlayan işlem olduğuna dair kanıt yoktur.

## Güvenli Geri Dönüş Noktası

Onarım başlamadan önce bozuk durumun tam kopyası alındı:

`<RECOVERY_BACKUP_DIR>`

- 267 dosya
- 13.682.628 bayt
- Kritik kaynak ve Git index hash'leri kaynakla birebir doğrulandı

Bu yedek, onarım öncesi adli durumun korunmuş kopyasıdır.

## Kurtarılan Üretim Dosyaları

| Dosya | Boyut | Satır | SHA-256 |
|---|---:|---:|---|
| `app.py` | 35.466 bayt | 985 | `64f878e42e64cdc37e798cee0dc7fe5437b38c0f3e4500857434bca55cd2f8cf` |
| `extract_and_compare.py` | 94.267 bayt | 2.038 | `82e477e09f7c1279342397c1685de15946d74247665b3645d930cd852cd466c7` |
| `new_mapping_parser.py` | 10.982 bayt | 320 | `c7f64a31cea330b3e138b6b170cb5a7e798415273ed2577da484c715d2b65403` |
| `PROJE_DOKUMANTASYONU.md` | 87.943 bayt | 1.676 | `73fb561d866060de80ef32c4af014ca241f5c1698d86b52ecb6cb7d8bb44d110` |

Dokümantasyon ve çekirdek, 12 Ağustos Codex oturum kayıtlarından satır satır yeniden kuruldu. Mapping parser için yerel geçmişteki eski sürüm tek başına doğru değildi; 11 Ağustos kuralları geri uygulanıp gerçek veri paritesi ve canonical hash ile doğrulandı. `app.py`, son sağlam sürümdeki gibi doğrudan golden `extract_and_compare.py` çekirdeğine bağlandı.

## Geri Getirilen Tarihsel Yardımcılar

VS Code Local History ve Qwen file-history kayıtlarında birebir byte kopyası bulunan 13 yardımcı dosya geri getirildi:

- `analyze_region.py`
- `analyze_updated_mapping.py`
- `compare_extraction_v2.py`
- `compare_files.py`
- `compare_regions.py`
- `debug_analysis.py`
- `deep_analysis.py`
- `diagnostic_compare.py`
- `extract_vangolu_headers.py`
- `extract_vangolu_headers2.py`
- `find_dagitim_bedeli.py`
- `find_discrepancies.py`
- `fix_mapping.py`

Bu 13 dosyanın hedefteki SHA-256 değerleri, seçilen tarihsel kopyalarla 13/13 eşleşti. Bunlar eski tanı/araştırma araçlarıdır; üretim uygulaması tarafından import edilmezler. Bazılarında taşınma öncesi sabit klasör yolları bulunduğu için bağımsız çalıştırılmadan önce ayrıca modernize edilmeleri gerekebilir.

## Kaynağı Bulunamayan 31 Yardımcı Dosya

Aşağıdaki dosyalar 0 bayttır ve taranan Git, VS Code, Qwen, Codex ve ZIP kaynaklarında güvenilir içerik kopyası bulunamadı:

- `analyze_bogazici_terim.py`
- `analyze_remaining_823_differences.py`
- `analyze_toroslar_20_mm.py`
- `compare_1005539_1005548.py`
- `compare_bogazici_ek_src_ref.py`
- `compare_toroslar_all.py`
- `compare_trakya_detailed.py`
- `comprehensive_diagnosis.py`
- `count_bogazici_ref_tariffs.py`
- `debug_etso_format.py`
- `debug_html_parser.py`
- `debug_notlar5_ext.py`
- `debug_osmangazi_guc_step_by_step.py`
- `debug_osmangazi_headers.py`
- `debug_osmangazi_headers_norm.py`
- `debug_osmangazi_mapping_keys.py`
- `debug_osmangazi_read.py`
- `debug_sources_v4.py`
- `debug_xml_cells.py`
- `deep_bogazici_analysis.py`
- `deep_debug_notlar5.py`
- `deep_debug_notlar5_v3.py`
- `deep_investigate_mismatch.py`
- `diagnose_4_missing_rows.py`
- `diagnostic_quick.py`
- `export_final_combined_excel.py`
- `final_field_eval.py`
- `find_osmangazi_etso.py`
- `find_toroslar_120160_xml.py`
- `find_toroslar_threshold.py`
- `notlar2_full_diagnosis.py`

Bu 31 dosyanın hiçbiri üretim kodu tarafından import edilmiyor. İçerikleri tahmin edilerek yeniden yazılmadı; böylece yanlış tarihsel kod projeye sokulmadı.

## Doğrulama Sonuçları

### Çekirdek ve Veri Paritesi

- Beklenen kayıt: 2.593; gerçekleşen: 2.593
- Beklenen bölge: 21; gerçekleşen: 21
- Beklenen toplam: 353.348.178,79 TL; gerçekleşen: 353.348.178,79 TL
- Beklenen canonical SHA-256 ile birebir eşleşme
- Vangölü alt klasörleri dahil 9 kayıt
- 36 kaynak dosya keşfedildi, filtreler sonrası 30 dosya işlendi
- 3.374 ham kayıt, 774 birleştirilmiş mükerrer kayıt

### Derin Dosya Taraması

Üç seviyeli geçici klasör testinde `.XLS`, `.XlSx`, `.XML` ve `.HtMl` dosyalarının tamamı bulundu. `.txt` ve `~$` geçici Excel kilit dosyaları dışlandı.

### Türkçe Sayı ve ETSO

Türkçe/ABD sayı biçimleri, parantezli ve sonda eksi işaretli değerler doğru sayıya çevrildi. ETSO normalizasyon örnekleri geçti. Arayüz biçimi `1.000.000,34 TL`; ETSO görünümü `12***678` olarak doğrulandı.

### Excel Çıktısı

- 2.593 veri satırı ve 17 sütun
- İndirilen Excel'de ETSO maskesiz ve kaynakla aynı
- O ve P formülleri 2.593/2.593 satırda mevcut
- 17/17 sütun genişliği `max_length + 2` ile birebir

### Streamlit

- İlk açılış: 0 exception, 0 UI error
- Gerçek analiz butonu: 0 exception, 0 UI error
- Önizleme: 2.593 satır; 99 satır sınırı yok
- Ekranda ETSO maskeli, indirilen Excel'de maskesiz
- Normal rerun sonrası DataFrame, istatistikler ve Excel bytes korunuyor
- 9.999.999,99 TL kayıt tutar uyarısına girmiyor
- 10.000.000,00 TL eşik kaydı Türkçe format ve maskeli ETSO ile uyarıya giriyor
- AI context yalnız toplu istatistik ve şema içeriyor; müşteri/ETSO satır değerleri sızmıyor
- Mock Gemini timeout durumunda uygulama çökmüyor ve `Bağlantı kurulamadı: ...` mesajı gösteriyor

### Çalışma Zamanı

- 64 Python dosyası AST kontrolünde 0 sözdizimi hatası
- Üretim modülleri import testi başarılı
- `pip check`: bağımlılık hatası yok
- Streamlit health: HTTP 200 / `ok`
- Ana sayfa: HTTP 200
- Sunucu yalnız `127.0.0.1` adresine bağlanıyor

## 17 Ağustos Güvenlik Devamı

Kullanıcı onayıyla kurtarma sonrasında aşağıdaki güvenlik temizliği uygulandı:

- Eski Gemini anahtarı `app.py` ve proje dokümantasyonundan kaldırıldı; çalışma ağacında eski anahtar eşleşmesi sıfıra indirildi.
- Gemini anahtarı tembel yüklenen `GEMINI_API_KEY` ortam değişkeni veya korumalı `st.secrets` üzerinden okunur hale getirildi.
- Secret bulunmaması uygulamanın veri işleme ekranını engellemez; yalnız chatbot açıklayıcı bir yapılandırma hatası döndürür.
- Hata mesajlarında kullanılan anahtar otomatik olarak `[API anahtarı gizlendi]` metniyle redakte edilir.
- Kök `.gitignore` eklendi; `.env`, Streamlit secrets, Python bytecode/cache ve özel anahtar dosyaları sürüm takibi dışında bırakıldı.
- Git tarafından izlenen 13 `.pyc` dosyası ve test sırasında oluşan parser bytecode'u fiziksel olarak kaldırıldı.
- Güvenli `secrets.toml.example` ve çalıştırma/yapılandırma açıklamaları eklendi.
- Her iki launcher Streamlit'i uygulama klasörünü çalışma dizini yaparak ve yalnız `127.0.0.1` üzerinde başlatacak şekilde doğrulandı.
- Doğrudan kullanılan `numpy` bağımlılığı `requirements.txt` içine açıkça eklendi.
- Güvenlik değişikliklerinden sonra golden veri regresyonu tekrar geçti: 2.593 kayıt, 21/21 bölge, 353.348.178,79 TL ve beklenen canonical SHA-256.
- Post-security Streamlit AppTest gerçek analiz, rerun/state, Excel, secret bulunmama ve mock Gemini başarı/timeout/redaksiyon senaryolarında 0 exception ve 0 UI error ile geçti.
- Yerel HTTP smoke testinde health ve ana sayfa HTTP 200 döndü; test sunucusu doğrulama sonunda kapatıldı.

## Açık Riskler ve Sonraki Adımlar

1. GitHub deposu herkese açıktır ve eski Gemini anahtarı public `ede7f45` commit geçmişinde bulunur. Çalışma ağacı temizlenmiş olsa da anahtar Google tarafında mutlaka iptal edilip yenilenmelidir.
2. Adli kurtarma yedeği eski anahtarlı `app.py` ve `app.pyc` kopyalarını içerir. Yedek erişimi kısıtlı/offline tutulmalı; anahtar rotasyonundan sonra ayrıca scrub veya güvenli imha kararı verilmelidir.
3. Normal push eski commit'i geçmişte bırakır. Geçmişi temizlemek için mevcut public `main` dalını değiştirecek kontrollü bir history rewrite gerekir; bu işlem kullanıcı tercihi ve `--force-with-lease` koruması olmadan yapılmamalıdır.
4. Eski `google.generativeai` SDK'sı kullanım ömrü sonu uyarısı veriyor. Ayrı ve testli bir değişiklikte güncel Google GenAI SDK'sına geçilmelidir.
5. `st.components.v1.html` mevcut Streamlit sürümünde çalışıyor ancak deprecation uyarısı veriyor. Kopyalama kısayolu davranışı korunarak desteklenen alternatife geçiş araştırılmalıdır.
6. Kaynağı bulunamayan 31 tanı dosyası üretime bağlı değildir. İhtiyaç doğarsa eski içeriği tahmin etmek yerine görev bazlı yeni tanı araçları yazılmalıdır.

## Git Durumu

İlk kurtarma raporu hazırlanırken hiçbir commit veya push yapılmadı. Kullanıcı devam onayından sonra kurtarma ve güvenlik değişiklikleri yerel `main` dalında `Recover golden pipeline and secure Streamlit app` commit'i altında korumaya alındı. `origin/main` değiştirilmedi; public uzak depo anahtar rotasyonu ve normal push/history rewrite tercihi netleşmeden güncellenmedi.
