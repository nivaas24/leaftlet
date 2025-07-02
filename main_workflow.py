"""
PDF Reconstruction Workflow
Complete workflow for analyzing, extracting, modifying, and reconstructing PDFs
"""

import sys
import os
sys.path.append('src')

from pdf_analyzer import PDFAnalyzer
from pdf_to_word import PDFToWordConverter
from word_to_pdf import WordToPDFConverter

def main():
    print("=== PDF Reconstruction Workflow ===\n")
    
    # Step 1: Generate sample PDF (if it doesn't exist)
    if not os.path.exists("samples/sample.pdf"):
        print("Step 1: Generating sample PDF...")
        os.system("python src/generate_sample_pdf.py")
        print("✓ Sample PDF created\n")
    else:
        print("Step 1: Sample PDF already exists\n")
    
    # Step 2: Analyze PDF structure
    print("Step 2: Analyzing PDF structure...")
    with PDFAnalyzer("samples/sample.pdf") as analyzer:
        analyzer.save_analysis("output/analysis.json")
        analyzer.extract_images_to_folder("output/images")
    print("✓ PDF analysis complete\n")
    
    # Step 3: Convert PDF to Word
    print("Step 3: Converting PDF to Word document...")
    pdf_to_word_converter = PDFToWordConverter("samples/sample.pdf", "output/extracted.docx")
    pdf_to_word_converter.convert_to_word()
    print("✓ Word document created\n")
    
    # Step 4: Instructions for user
    print("Step 4: Manual modification required")
    print("- Open 'output/extracted.docx' in Microsoft Word")
    print("- Make your desired changes (add/remove/replace text)")
    print("- Save the file as 'output/modified.docx'")
    print("- Then run: python src/word_to_pdf.py")
    print("\nAlternatively, you can proceed with the current Word document as-is:\n")
    
    # Step 5: Optional automatic conversion back to PDF
    user_input = input("Would you like to convert the extracted Word document back to PDF now? (y/n): ")
    if user_input.lower() == 'y':
        print("\nStep 5: Converting Word back to PDF...")
        try:
            # Copy the extracted file to modified for the conversion
            import shutil
            shutil.copy("output/extracted.docx", "output/modified.docx")
            
            word_to_pdf_converter = WordToPDFConverter("output/modified.docx", "output/reconstructed.pdf")
            word_to_pdf_converter.convert_to_pdf()
            print("✓ PDF reconstruction complete")
            print("✓ Check 'output/reconstructed.pdf' for the result")
        except Exception as e:
            print(f"Error during conversion: {e}")
            print("Please ensure Microsoft Word is installed for docx2pdf to work properly")
    
    print("\n=== Workflow Summary ===")
    print("Files created:")
    print("- samples/sample.pdf (original PDF)")
    print("- output/analysis.json (structural analysis)")
    print("- output/extracted.docx (extracted content)")
    print("- output/images/ (extracted images)")
    if os.path.exists("output/reconstructed.pdf"):
        print("- output/reconstructed.pdf (reconstructed PDF)")
    
    print("\n=== Next Steps ===")
    print("1. Edit 'output/extracted.docx' with your desired changes")
    print("2. Save it as 'output/modified.docx'")
    print("3. Run: python src/word_to_pdf.py")
    print("4. Your reconstructed PDF will be saved as 'output/final.pdf'")

if __name__ == "__main__":
    main()
