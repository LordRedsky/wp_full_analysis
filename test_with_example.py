import io
from analysis_sert import process_certificate_analysis

def test_with_example_file():
    print("Testing certificate analysis with sert example.xlsx...")
    
    # Read the example file
    with open("sert example.xlsx", "rb") as f:
        sertifikat_bytes = f.read()
    
    # Process the certificate analysis
    result_bytes = process_certificate_analysis(sertifikat_bytes)
    
    # Save the result to a new file
    with open("sertifikat_analysed.xlsx", "wb") as f:
        f.write(result_bytes)
    
    print("Certificate analysis completed successfully!")
    print("Output saved as 'sertifikat_analysed.xlsx'")

if __name__ == "__main__":
    test_with_example_file()