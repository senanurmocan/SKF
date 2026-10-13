"""
Integrated AI Data Assistant Module for EDAS Bill Analytics using Google Gemini (google.generativeai).
"""

import google.generativeai as genai
from core.normalizer import mask_etso


def generate_data_context(extracted_data, stats, audit_results):
    """Generates structured context summary from processed dataset."""
    if not extracted_data:
        return "Veri seti henüz yüklenmedi."

    total_records = len(extracted_data)
    read_region_count = stats.get('read_region_count', 0) if stats else 0
    successful_regs = stats.get('successful_regions', []) if stats else []
    missing_regs = stats.get('missing_regions', []) if stats else []

    region_counts = {}
    max_bill_val = 0.0
    max_bill_row = None

    for r in extracted_data:
        reg = r.get('Dağıtım Bölgesi', 'Bilinmeyen')
        region_counts[reg] = region_counts.get(reg, 0) + 1

        dagitim = float(r.get('Dağıtım Bedeli(TL)', 0) or 0)
        guc = float(r.get('Güç Bedeli(TL)', 0) or 0)
        toplam = float(r.get('Toplam (TL)', 0) or (dagitim + guc))
        val = abs(toplam if toplam > 0 else dagitim)

        if val > max_bill_val and val < 50_000_000_000.0:
            max_bill_val = val
            max_bill_row = r

    max_bill_info = "Yok"
    if max_bill_row:
        m_etso = mask_etso(max_bill_row.get('Etso Kodu', ''))
        m_reg = max_bill_row.get('Dağıtım Bölgesi', '')
        m_mus = max_bill_row.get('Müşteri', 'Bilinmeyen Abone')
        max_bill_info = f"{max_bill_val:,.2f} TL - Bölge: {m_reg}, ETSO: {m_etso}, Müşteri: {m_mus}"

    context = f"""
=== CK ENERJİ EDAŞ FATURA VERİ SETİ ÖZETİ ===
- Toplam Çıkarılan Kayıt Sayısı: {total_records:,}
- Okunan Eşsiz Bölge Sayısı: {read_region_count} / 21 EDAŞ Bölgesi
- Okunan Bölgeler Listesi: {', '.join(successful_regs)}
- Eksik / Okunamayan Bölgeler: {', '.join(missing_regs) if missing_regs else 'Yok (21/21 Tam Uyum)'}
- En Yüksek Fatura Kaydı: {max_bill_info}
- Veri Sağlık Skoru: %{audit_results.get('health_score', 100) if audit_results else 100}

=== BÖLGE BAZLI SATIR SAYILARI ===
"""
    for reg_name, count in sorted(region_counts.items(), key=lambda x: x[1], reverse=True):
        context += f"- {reg_name}: {count:,} satır\n"

    if audit_results and audit_results.get('warnings'):
        context += "\n=== ANOMALİ VE UYARI LOGLARI ===\n"
        for w in audit_results['warnings'][:15]:
            context += f"- [{w['category']}] {w['message']}\n"

    return context


def answer_user_query(query, extracted_data, stats, audit_results, api_key=None):
    """
    Answers user question using Google Gemini AI (google.generativeai SDK).
    """
    if not api_key or not str(api_key).strip():
        return "⚠️ Lütfen asistanı kullanmak için sidebar'dan Gemini API Key giriniz."

    if not extracted_data:
        return "⚠️ Henüz bir veri seti işlenmedi. Lütfen önce soldaki '🚀 Verileri Birleştir ve Analiz Et' butonuna tıklayarak verileri yükleyin."

    context = generate_data_context(extracted_data, stats, audit_results)
    prompt = f"""Sen CK Enerji EDAŞ Fatura Veri Analiz Asistanısın. Kullanıcının sorularına aşağıdaki gerçek veri setinin özetine tam sadık kalarak Türkçe ve net cevap ver.

GERÇEK VERİ SETİ CONTEXTİ:
{context}

KULLANICI SORUSU:
{query}
"""

    try:
        genai.configure(api_key=str(api_key).strip())
        try:
            model = genai.GenerativeModel("gemini-1.5-flash")
            response = model.generate_content(prompt)
            return response.text
        except Exception:
            model = genai.GenerativeModel("gemini-pro")
            response = model.generate_content(prompt)
            return response.text
    except Exception as e:
        return f"❌ Google Gemini AI Bağlantı Hatası: {str(e)}"
