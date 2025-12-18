import pandas as pd
import json
import re
from datetime import datetime

def clean_text(text):
    """Clean and normalize text"""
    if pd.isna(text):
        return ''
    return str(text).strip().replace('\n', ' ').replace('  ', ' ')

def extract_inn(text):
    """Extract INN from combined fields"""
    if pd.isna(text):
        return ''
    text = str(text)
    # Look for INN patterns (10 or 12 digits)
    inn_match = re.search(r'\b(\d{10}|\d{12})\b', text)
    if inn_match:
        return inn_match.group(1)
    return text.strip()

def determine_member_type(name):
    """Determine member type from name"""
    if pd.isna(name):
        return 'unknown'
    
    name_lower = str(name).lower()
    
    if 'ип ' in name_lower or 'индивидуальный предприниматель' in name_lower:
        return 'ie'
    elif any(word in name_lower for word in ['ооо', 'общество с ограниченной ответственностью', 
                                               'limited', 'ltd', 'corporation', 'акционерное общество', 'ао ']):
        return 'company'
    else:
        return 'person'

def create_short_name(full_name, member_type):
    """Create short name based on member type"""
    if pd.isna(full_name):
        return ''
    
    name = str(full_name).strip()
    
    if member_type == 'company':
        # Extract company name in quotes or after ООО
        if '«' in name and '»' in name:
            short = name[name.find('«'):name.find('»')+1]
            if 'ООО' in name[:30]:
                return f"ООО {short}"
            return short
        elif '"' in name:
            parts = name.split('"')
            if len(parts) >= 2:
                return f'ООО "{parts[1]}"' if 'ООО' in name else f'"{parts[1]}"'
        # Fallback
        return name[:60]
    
    elif member_type == 'ie':
        # For IE, extract name after ИП
        if 'ИП ' in name:
            name_part = name.split('ИП ')[-1].strip()
            parts = name_part.split()
            if len(parts) >= 3:
                return f"ИП {parts[0]} {parts[1][0]}.{parts[2][0]}."
        return name[:60]
    
    else:  # person
        # For persons, keep full name
        return name

def is_separator_row(row):
    """Check if row is a section separator (excluded members header)"""
    first_col = str(row.iloc[0]).lower() if not pd.isna(row.iloc[0]) else ''
    return 'исключенные' in first_col or 'добровольно' in first_col or 'выбывшие' in first_col

def process_excel_to_json(excel_file):
    """Process Excel file with proper structure detection"""
    
    # Read the entire Excel file
    df_raw = pd.read_excel(excel_file, header=None)
    
    print(f"Total rows in Excel: {len(df_raw)}")
    
    # Find where the actual data starts (after title and header rows)
    data_start = None
    for idx, row in df_raw.iterrows():
        first_cell = str(row.iloc[0]).strip()
        if first_cell.isdigit() and int(first_cell) == 1:
            data_start = idx
            break
    
    if data_start is None:
        print("ERROR: Could not find data start row")
        return None
    
    print(f"Data starts at row: {data_start}")
    
    # Find separator between active and excluded members
    separator_row = None
    for idx in range(data_start + 1, len(df_raw)):
        if is_separator_row(df_raw.iloc[idx]):
            separator_row = idx
            break
    
    print(f"Separator row (excluded members start): {separator_row}")
    
    # Process active members
    active_df = df_raw.iloc[data_start:separator_row] if separator_row else df_raw.iloc[data_start:]
    
    # Process excluded members (if separator found)
    excluded_df = df_raw.iloc[separator_row+1:] if separator_row and separator_row+1 < len(df_raw) else None
    
    all_members = []
    
    # Process active members
    for idx, row in active_df.iterrows():
        # Skip if first column is not a number
        if pd.isna(row.iloc[0]) or not str(row.iloc[0]).strip().isdigit():
            continue
        
        name = clean_text(row.iloc[1])
        if not name:
            continue
        
        member_type = determine_member_type(name)
        
        member = {
            'id': str(int(float(row.iloc[0]))),
            'name': name,
            'shortName': create_short_name(name, member_type),
            'inn': extract_inn(row.iloc[2]) if len(row) > 2 else '',
            'address': clean_text(row.iloc[3]) if len(row) > 3 else '',
            'activity': clean_text(row.iloc[4]) if len(row) > 4 else '',
            'interests': clean_text(row.iloc[5]) if len(row) > 5 else '',
            'type': member_type,
            'status': 'active',
            'statusText': 'Действующий член НКО ПОВС "ПАРУС"',
            'dateJoined': ''
        }
        
        all_members.append(member)
    
    # Process excluded members
    if excluded_df is not None and len(excluded_df) > 0:
        for idx, row in excluded_df.iterrows():
            # Skip header rows and empty rows
            if pd.isna(row.iloc[0]) or not str(row.iloc[0]).strip().isdigit():
                continue
            
            name = clean_text(row.iloc[1])
            if not name:
                continue
            
            member_type = determine_member_type(name)
            
            member = {
                'id': str(int(float(row.iloc[0]))),
                'name': name,
                'shortName': create_short_name(name, member_type),
                'inn': extract_inn(row.iloc[2]) if len(row) > 2 else '',
                'address': clean_text(row.iloc[3]) if len(row) > 3 else '',
                'activity': clean_text(row.iloc[4]) if len(row) > 4 else '',
                'interests': clean_text(row.iloc[5]) if len(row) > 5 else '',
                'type': member_type,
                'status': 'inactive',
                'statusText': 'Исключен из членов НКО ПОВС "Парус"',
                'dateExcluded': ''
            }
            
            all_members.append(member)
    
    # Separate by type and status
    active_companies = [m for m in all_members if m['status'] == 'active' and m['type'] == 'company']
    active_ie = [m for m in all_members if m['status'] == 'active' and m['type'] == 'ie']
    active_persons = [m for m in all_members if m['status'] == 'active' and m['type'] == 'person']
    
    inactive_companies = [m for m in all_members if m['status'] == 'inactive' and m['type'] == 'company']
    inactive_ie = [m for m in all_members if m['status'] == 'inactive' and m['type'] == 'ie']
    inactive_persons = [m for m in all_members if m['status'] == 'inactive' and m['type'] == 'person']
    
    # Create comprehensive output
    output = {
        'metadata': {
            'totalMembers': len(all_members),
            'activeMembers': len([m for m in all_members if m['status'] == 'active']),
            'inactiveMembers': len([m for m in all_members if m['status'] == 'inactive']),
            'companies': len([m for m in all_members if m['type'] == 'company']),
            'individualEntrepreneurs': len([m for m in all_members if m['type'] == 'ie']),
            'privatePersons': len([m for m in all_members if m['type'] == 'person']),
            'lastUpdated': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'sourceFile': excel_file
        },
        'members': all_members
    }
    
    # Save complete file
    with open('members_complete.json', 'w', encoding='utf-8') as f:
        json.dump(output, f, ensure_ascii=False, indent=2)
    
    # Save compact search version (only essential fields)
    search_data = []
    for m in all_members:
        search_data.append({
            'id': m['id'],
            'name': m['name'],
            'shortName': m['shortName'],
            'inn': m['inn'],
            'type': m['type'],
            'status': m['status'],
            'statusText': m['statusText']
        })
    
    with open('members_search.json', 'w', encoding='utf-8') as f:
        json.dump(search_data, f, ensure_ascii=False, indent=2)
    
    # Statistics
    print("\n" + "="*60)
    print("PROCESSING COMPLETE")
    print("="*60)
    print(f"\nTotal Members: {len(all_members)}")
    print(f"  Active: {len([m for m in all_members if m['status'] == 'active'])}")
    print(f"  Inactive: {len([m for m in all_members if m['status'] == 'inactive'])}")
    print(f"\nBy Type:")
    print(f"  Companies: {len([m for m in all_members if m['type'] == 'company'])}")
    print(f"    - Active: {len(active_companies)}")
    print(f"    - Inactive: {len(inactive_companies)}")
    print(f"  Individual Entrepreneurs: {len([m for m in all_members if m['type'] == 'ie'])}")
    print(f"    - Active: {len(active_ie)}")
    print(f"    - Inactive: {len(inactive_ie)}")
    print(f"  Private Persons: {len([m for m in all_members if m['type'] == 'person'])}")
    print(f"    - Active: {len(active_persons)}")
    print(f"    - Inactive: {len(inactive_persons)}")
    print(f"\nGenerated Files:")
    print(f"  ✓ members_complete.json ({len(all_members)} members with all fields)")
    print(f"  ✓ members_search.json (compact version for fast search)")
    print("="*60)
    
    return output

# Run
if __name__ == "__main__":
    import sys
    
    # Get filename from command line or use default
    filename = sys.argv[1] if len(sys.argv) > 1 else "Сведения о членах НКО ПОВС Парус на 01.12.2025.xls"
    
    print(f"Processing: {filename}\n")
    result = process_excel_to_json(filename)
    
    if result:
        print("\n✅ SUCCESS! Upload members_complete.json or members_search.json to your hosting.")
    else:
        print("\n❌ FAILED! Check the error messages above.")