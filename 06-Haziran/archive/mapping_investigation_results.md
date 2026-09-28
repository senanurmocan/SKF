# Dağıtım Bedeli Column Mapping Investigation Results

## Investigation Summary
Root cause: Column mapping file (SKF Başlıkları Güncel.xlsx) has incorrect column letters for several regions, causing extraction to read wrong columns with extreme values.

## Findings

### CORRECT Mappings (Verified)
| Region | Column Letter | Index | Status |
|--------|---------------|-------|--------|
| AKEDAŞ | X | 23 | ✓ MATCH |
| Akdeniz EDAŞ | AU | 46 | ✓ MATCH |
| Boğaziçi EDAŞ | BA | 52 | ✓ MATCH |
| Trakya EDAŞ | AT | 45 | ✓ MATCH |
| Çamlıbel EDAŞ | P | 15 | ✓ MATCH |

### INCORRECT Mappings (Mismatches)
| Region | Expected | Found | Offset | Issue |
|--------|----------|-------|--------|-------|
| Kayseri ve Civarı | Y (24) | 1 | -23 | Wrong column letter in mapping |
| Meram EDAŞ | AN (39) | 38 | -1 | Off by 1 |
| Osmangazi EDAŞ | BT (71) | 14 | -57 | Major mismatch - mapping likely wrong |
| Sakarya EDAŞ | AS (44) | 43 | -1 | Off by 1 |

### NO EXCEL FILES FOUND
| Region | Folder Name | Status |
|--------|-------------|--------|
| ADM EDAŞ | ADM EDAŞ | No Excel files |
| AYEDAŞ | AYEDAŞ | No Excel files |
| Aras EDAŞ | Aras EDAŞ | No Excel files |
| Başkent EDAŞ | Başkent EDAŞ | No Excel files |
| Fırat EDAŞ | Fırat EDAŞ | No Excel files |
| Gediz EDAŞ | Gediz EDAŞ | No Excel files |
| Toroslar EDAŞ | Toroslar EDAŞ | No Excel files |
| Uludağ EDAŞ | Uludağ EDAŞ | No Excel files |
| Vangölü EDAŞ | Vangölü EDAŞ | No Excel files |
| Yeşilırmak EDAŞ | Yeşilırmak EDAŞ | No Excel files |
| Çoruh EDAŞ | Çoruh EDAŞ | No Excel files |

### PROBLEMATIC FILES
| Region | File | Issue |
|--------|------|-------|
| Dicle EDAŞ | f47_9584715_20260707_2031.xlsx | Corrupted/特殊 format - headers are merge cells with title text, no proper column headers |

## Root Cause Analysis
1. **Mapping file column letters incorrect**: For Kayseri, Osmangazi, Meram, Sakarya, mapping file has wrong column letters
2. **Missing Excel files**: 11 regions show "NO EXCEL FILES" - either files don't exist or folders are empty
3. **Corrupted file**: Dicle EDAŞ file has special format that can't be parsed by standard Excel readers

## Impact on Extreme Values
- **Original finding**: Mapping file says "AU" (index 46) for Akdeniz, actual file has index 46 ✓ - this was CORRECT
- **But**: Osmangazi mapping says BT (71) but file only has 14 columns - reading way off
- **And**: Kayseri mapping says Y (24) but file has only 1 column with Dağıtım Bedeli

## Recommendations
1. **Update mapping file** with correct column letters for:
   - Kayseri ve Civarı: Change from Y to actual column containing "Dağıtım Bedeli"
   - Meram EDAŞ: Change from AN to AM (index 38)
   - Osmangazi EDAŞ: Change from BT to actual column
   - Sakarya EDAŞ: Change from AS to AR (index 43)

2. **Verify missing region folders** - check if Excel files exist elsewhere or have different names

3. **Handle Dicle EDAŞ** - manually inspect file structure or request properly formatted file

## Files Modified
- `verify_column_mapping_v2.py` - Fixed region name matching logic
- `find_dagitim_bedeli.py` - New script for comprehensive column mapping verification
- `check_dicle.py` - Debug script for Dicle EDAŞ
