"""
Format Tracker & Dynamic Header Adaptation Monitor.
Detects column header changes in source files and confirms dynamic adaptation.
"""


class FormatTracker:
    def __init__(self):
        self.adaptations = []
        self.unmatched_headers = []

    def record_adaptation(self, region_name, original_header, matched_synonym, column_letter):
        """Record when dynamic header resolution adapted to a non-standard column name."""
        msg = f"💡 ADAPTASYON: {region_name} bölgesinde '{original_header}' başlığı dinamik olarak '{matched_synonym}' (Sütun {column_letter}) ile eşleştirildi."
        self.adaptations.append({
            'region': region_name,
            'header': original_header,
            'matched': matched_synonym,
            'col': column_letter,
            'message': msg
        })

    def record_unmatched(self, region_name, field_name):
        """Record when a field could not be matched for a region."""
        msg = f"⚠️ DİKKAT: {region_name} bölgesinde '{field_name}' başlığı kaynak dosyada bulunamadı (Varsayılan 0 alındı)."
        self.unmatched_headers.append({
            'region': region_name,
            'field': field_name,
            'message': msg
        })

    def get_summary_messages(self):
        """Get list of user-facing UI notifications."""
        messages = []
        for adapt in self.adaptations:
            messages.append(adapt['message'])
        for unmatch in self.unmatched_headers:
            messages.append(unmatch['message'])
        return messages
