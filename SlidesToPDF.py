import os
import sys
import re
from pathlib import Path

try:
    import aspose.slides as slides
    import fitz
except ImportError:
    print("Error: Missing dependencies.")
    print("Run: pip install aspose-slides pymupdf")
    input("\nPress Enter to exit...")
    sys.exit(1)

def surgical_clean_watermark(pdf_path):
    """Removes watermark streams and updates document metadata."""
    try:
        doc = fitz.open(pdf_path)
        
        # Set PDF title to filename to avoid default browser tab titles
        metadata = doc.metadata
        metadata["title"] = Path(pdf_path).stem
        doc.set_metadata(metadata)
        
        triggers = [
            b"Aspose.Slides", 
            b"Evaluation only", 
            b"Created with", 
            b"Python via .NET", 
            b"Copyright 2004-2026",
            b"Aspose Pty Ltd"
        ]

        for page in doc:
            contents = page.get_contents()
            for xref in contents:
                stream = doc.xref_stream(xref)
                new_stream = stream
                
                for match in re.finditer(b"BT.*?ET", stream, re.DOTALL):
                    block = match.group(0)
                    if any(t in block for t in triggers):
                        new_stream = new_stream.replace(block, b"")
                
                if new_stream != stream:
                    doc.update_stream(xref, new_stream)
        
        temp_path = pdf_path.with_name(f"fixed_{pdf_path.name}")
        doc.save(temp_path, garbage=4, deflate=True)
        doc.close()
        
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        os.rename(temp_path, pdf_path)
        return True
    except Exception as e:
        print(f"Cleaning failed: {e}")
        return False

def convert_to_pdf(input_path):
    input_path = input_path.strip('"')
    if not os.path.exists(input_path):
        return

    file_path = Path(input_path)
    if file_path.suffix.lower() not in ['.ppt', '.pptx', '.odp']:
        return

    print(f"\nProcessing: {file_path.name}")
    pdf_path = file_path.with_suffix(".pdf")

    try:
        with slides.Presentation(str(file_path)) as presentation:
            presentation.save(str(pdf_path), slides.export.SaveFormat.PDF)
        
        surgical_clean_watermark(pdf_path)
        print(f"==> Success: {pdf_path.name}")
    except Exception as e:
        print(f"==> Failed: {e}")

if __name__ == "__main__":
    print("========================================")
    print("   Slides to PDF Converter")
    print("========================================")

    if len(sys.argv) > 1:
        for arg in sys.argv[1:]:
            convert_to_pdf(arg)
    else:
        print("\nUsage: Drag and drop files onto 'Convert.bat'.")

    print("\n========================================")
    input("Done. Press Enter to close...")
