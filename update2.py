import fitz  # PyMuPDF
import os
from datetime import datetime

def fill_dpr_template_smart(input_pdf_path, name, address):
    """
    DPR Template PDF లో Table 3, row (i) "Name and address" cell లో fill చేస్తుంది
    Address పెద్దదైతే cell size automatically adjust అవుతుంది
    """
    
    # Output path
    downloads_path = os.path.join(os.path.expanduser('~'), 'Downloads')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"DPR_Template_Filled_{timestamp}.pdf"
    output_pdf_path = os.path.join(downloads_path, output_filename)
    
    try:
        # Original PDF open చేయడం
        doc = fitz.open(input_pdf_path)
        
        # Find the page with "Information about SPV"
        target_page = None
        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text()
            if "Information about SPV" in text or "Name and address" in text:
                target_page = page_num
                break
        
        if target_page is None:
            print("⚠️  Using default page 2")
            target_page = 1
        
        page = doc[target_page]
        
        print(f"\n✅ Working on page {target_page + 1}")
        
        # Search for "Details/ Compliance" header to find the third column
        details_header = page.search_for("Details/ Compliance")
        
        if details_header:
            # Found the Details/Compliance column header
            header_rect = details_header[0]
            print(f"📍 Found 'Details/Compliance' column at: {header_rect}")
            
            # Now find "Name and address" row
            name_address_instances = page.search_for("Name and address")
            
            if name_address_instances:
                name_rect = name_address_instances[0]
                print(f"📍 Found 'Name and address' row at: {name_rect}")
                
                # Cell position: use column x from header, row y from "Name and address"
                cell_x0 = header_rect.x0 + 5  # Start from Details column with padding
                cell_y0 = name_rect.y0 - 2    # Same row as "Name and address"
                cell_x1 = header_rect.x1 - 5  # End at column boundary with padding
                
                # Calculate height based on text length
                combined_text = f"{name}\n{address}"
                lines = combined_text.split('\n')
                line_height = 12
                required_height = len(lines) * line_height + 10
                
                cell_y1 = cell_y0 + max(required_height, 25)
                
                cell_rect = fitz.Rect(cell_x0, cell_y0, cell_x1, cell_y1)
            else:
                print("⚠️  Using column position only")
                cell_rect = fitz.Rect(header_rect.x0 + 5, 330, header_rect.x1 - 5, 360)
                combined_text = f"{name}\n{address}"
        else:
            print("⚠️  Headers not found, using default coordinates")
            # Fallback: Details/Compliance column typically starts around x=575
            cell_rect = fitz.Rect(575, 330, 840, 365)
            combined_text = f"{name}\n{address}"
        
        # Insert text in the cell
        rc = page.insert_textbox(
            cell_rect,
            combined_text,
            fontsize=9,
            fontname="helv",
            color=(0, 0, 0),
            align=fitz.TEXT_ALIGN_LEFT,
            overlay=True
        )
        
        if rc < 0:
            print("⚠️  Text might be too long for the cell, adjusting...")
            # If text doesn't fit, try with smaller font
            page.insert_textbox(
                cell_rect,
                combined_text,
                fontsize=8,
                fontname="helv",
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_LEFT,
                overlay=True
            )
        
        print(f"✅ Text inserted at: {cell_rect}")
        
        # Save PDF
        doc.save(output_pdf_path, garbage=4, deflate=True, clean=True)
        doc.close()
        
        return output_pdf_path, True
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
        if 'doc' in locals():
            doc.close()
        return None, False


def main():
    """
    Main function - Simple and clean workflow
    """
    print("\n" + "="*70)
    print(" "*15 + "DPR TEMPLATE PDF FILLER")
    print("="*70)
    
    # Step 1: Input PDF
    print("\n📄 Input PDF File")
    print("-" * 70)
    input_pdf = input("Enter PDF path (or press Enter for 'DPR Template.pdf'): ").strip()
    
    if not input_pdf:
        input_pdf = "DPR Template.pdf"
    
    # Clean up path
    input_pdf = input_pdf.strip('"').strip("'")
    
    # Check file
    if not os.path.exists(input_pdf):
        print(f"\n❌ ERROR: PDF not found!")
        print(f"   Looking for: {input_pdf}")
        print(f"\n💡 Make sure the PDF is in the same folder")
        print(f"   Or provide full path")
        return
    
    print(f"✅ Found: {os.path.basename(input_pdf)}")
    
    # Step 2: Get details
    print("\n👤 Enter Details")
    print("-" * 70)
    print("Format: Name, Address")
    print("Example: Kiran Kumar, 17/367-MTM, Guntur, Andhra Pradesh")
    user_input = input("\nEnter Name and Address (comma separated): ").strip()
    
    # Validate input
    if ',' not in user_input:
        print("\n❌ ERROR: Use comma to separate name and address")
        print("   Example: Kiran Kumar, 17/367-MTM")
        return
    
    # Split into name and address
    parts = user_input.split(',', 1)
    name = parts[0].strip()
    address = parts[1].strip() if len(parts) > 1 else ""
    
    if not name or not address:
        print("\n❌ ERROR: Both name and address are required")
        return
    
    print(f"\n✓ Name: {name}")
    print(f"✓ Address: {address}")
    
    # Step 3: Process
    print("\n⏳ Processing PDF...")
    print("-" * 70)
    
    output_path, success = fill_dpr_template_smart(input_pdf, name, address)
    
    # Step 4: Results
    if success and output_path and os.path.exists(output_path):
        file_size = os.path.getsize(output_path) / 1024
        
        print("\n" + "="*70)
        print(" "*25 + "✅ SUCCESS!")
        print("="*70)
        print(f"\n📄 File Name    : {os.path.basename(output_path)}")
        print(f"📁 Location     : {output_path}")
        print(f"📊 File Size    : {file_size:.2f} KB")
        print(f"\n👤 Name         : {name}")
        print(f"📍 Address      : {address}")
        print("\n" + "="*70)
        
        # Open Downloads folder
        print("\n📂 Opening Downloads folder...")
        try:
            os.startfile(os.path.dirname(output_path))
        except Exception as e:
            print(f"   Note: {e}")
        
        print("\n✨ PDF saved successfully!")
        print("="*70 + "\n")
        
    else:
        print("\n" + "="*70)
        print("❌ FAILED: Could not create PDF")
        print("="*70)
        print("\n💡 TROUBLESHOOTING:")
        print("   1. Install PyMuPDF: pip install PyMuPDF")
        print("   2. Check PDF is not corrupted")
        print("   3. Ensure write permissions to Downloads")
        print("="*70 + "\n")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Operation cancelled by user\n")
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}\n")
        import traceback
        traceback.print_exc()