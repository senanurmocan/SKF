def column_to_index(col_str):
    """Convert Excel column letter to 0-based index"""
    col_str = col_str.upper()
    result = 0
    for i, char in enumerate(reversed(col_str)):
        result += (ord(char) - ord('A') + 1) * (26 ** i)
    return result - 1

# Test
letters = ['AS', 'AR', 'AU', 'AT', 'BG', 'BT', 'AM', 'AN', 'B']
for letter in letters:
    print(f'{letter} = index {column_to_index(letter)}')
