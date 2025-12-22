from analysis_sert import _parse_names

# Test the specific case
result = _parse_names('A, B, C')
print(f'Input: A, B, C')
print(f'Parsed names: {result}')
print(f'Number of names: {len(result)}')

# Test other cases
test_cases = [
    'Siti, H. Ahmad, Drs. Budi, Ir. Candra',
    'H. Joko, A.Md.Kom, Siti, H. Ahmad',
    'Dr., Ir. Joko',
    'H. A. Budi',
    'La ilman, Wa ilma'
]

for test in test_cases:
    result = _parse_names(test)
    print(f'Input: {test}')
    print(f'  Parsed: {result} (count: {len(result)})')