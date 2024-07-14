import oletools.olevba

def extract_vba_macros(file_path):
    vbaparser = oletools.olevba.VBA_Parser(file_path)

    if vbaparser.detect_vba_macros():
        print(f"VBA Macros found in {file_path}:")
        for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
            print(f"\nFilename: {filename}")
            print(f"Stream Path: {stream_path}")
            print(f"VBA Filename: {vba_filename}")
            print(f"VBA Code:\n{vba_code}")
    else:
        print(f"No VBA Macros found in {file_path}")

    vbaparser.close()

# Example usage
file_path = r"C:\Users\harsh\Downloads\Download-Sample-File-xlsm.xlsm"
extract_vba_macros(file_path)
