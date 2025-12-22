#!/usr/bin/env python3
"""
Test the heirs detection feature with the example Excel file
"""

from analysis_sert import process_certificate_analysis

def test_with_example_file():
    """Test the analysis with the example file"""
    # Read the example file
    with open("sert example.xlsx", 'rb') as f:
        file_bytes = f.read()
    
    print("Processing example file with heirs detection...")
    result = process_certificate_analysis(file_bytes)
    
    print("Analysis completed successfully!")
    print(f"Data diproses: {result['data_diproses']}")
    print(f"NIB duplikat: {result['nib_duplikat']}")
    print(f"Jumlah Ahli Waris: {result['jumlah_ahli_waris']}")
    
    # Save the result
    with open("Sertifikat_analysed.xlsx", 'wb') as f:
        f.write(result['workbook_bytes'])
    
    print("Output saved to Sertifikat_analysed.xlsx")

if __name__ == "__main__":
    test_with_example_file()