import os
import json
from openai import OpenAI

# Import the original PDF generation function
try:
    from dpr_pdf_generator import create_dpr_pdf
except ImportError:
    print("Error: Please save the original 1300-line code as 'dpr_pdf_generator.py' in the same folder")
    exit(1)

# Initialize OpenAI client
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="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"
)

def get_llm_mapping(user_input):
    """
    Enhanced LLM mapping with support for custom text sections, paragraphs, and bullet points
    """
    
    system_prompt = """You are an expert DPR (Detailed Project Report) document analyzer. Your job is to extract information from natural language and map it to PDF fields INCLUDING custom text sections.

═══════════════════════════════════════════════════════════════
NEW CAPABILITY: CUSTOM TEXT SECTIONS
═══════════════════════════════════════════════════════════════

Users can now add CUSTOM TEXT under ANY section heading by using these patterns:

**Pattern 1: Section Text (Paragraphs under headings)**
User says: "Under introduction write: This project aims to revolutionize textile manufacturing"
→ Field: "Introduction_Text" = "This project aims to revolutionize textile manufacturing"

User says: "Add description under technical aspects: The technology will use Industry 4.0 principles"
→ Field: "Technical_Aspects_Text" = "The technology will use Industry 4.0 principles"

**Pattern 2: Bullet Points/Lists**
User says: "Objectives are: increase production, reduce costs, improve quality"
→ Field: "Objectives_List" = ["increase production", "reduce costs", "improve quality"]

User says: "Main benefits: 1) Higher efficiency 2) Better quality 3) Lower costs"
→ Field: "Benefits_List" = ["Higher efficiency", "Better quality", "Lower costs"]

**Pattern 3: Subsection Text**
User says: "Under raw materials mention: Cotton sourced locally, imported synthetic fibers"
→ Field: "Raw_materials_Text" = "Cotton sourced locally, imported synthetic fibers"

**Pattern 4: Multi-paragraph Text**
User says: "Add detailed description: First para about scope. Second para about benefits."
→ Field: "Detailed_Description" = "First para about scope.\n\nSecond para about benefits."

═══════════════════════════════════════════════════════════════
SECTION-SPECIFIC TEXT FIELDS (Map user input to these)
═══════════════════════════════════════════════════════════════

**Section 2.1 - Introduction Texts:**
- "Introduction_General_Text" → Text under "General scenario of industrial growth"
- "Introduction_Sector_Text" → Text under "Sector for which CFC is proposed"
- "Introduction_Cluster_Text" → Text under "Cluster and its products"
- "Introduction_Relevance_Text" → Text under "How CFC is relevant"

**Section 4 - Promoter Details Texts:**
- "Brief_bio_data_of_Promoters" → Biography text
- "Management_Setup_Text" → Management details text

**Section 6 - Implementation:**
- "Role_of_Implementing_Agency" → Agency role description

**Section 8 - Technical Aspects Texts:**
- "Scope_of_project_Text" → Project scope description
- "Locational_details_Text" → Location description
- "Technology_Text" → Technology description
- "Industry_4_0_Text" → Industry 4.0 details
- "Raw_materials_Text" → Raw materials description
- "Power_Text" → Power requirements
- "Water_Text" → Water requirements
- "Effluent_disposal_Text" → Effluent disposal details

**Section 10 - Project Components Texts:**
- "Land_Building_Details_Text" → Land and building description
- "Misc_fixed_assets" → Miscellaneous assets text
- "Preliminary_expenses_Text" → Preliminary expenses description
- "Pre-operative_expenses" → Pre-operative expenses text
- "Contingency_Provisions" → Contingency details
- "Working_Capital_Text" → Working capital description
- "Usage_Charges_Text" → How facilities will be used

**Section 11 - Fund Requirement:**
- "Fund_requirement_analysis_Text" → Fund availability analysis

**Section 12 - Usage Charges:**
- "Usage_Charges_Details" → Detailed usage charges

**Section 13 - Commercial Viability:**
- "Commercial_viability_comments" → Viability comments

**Section 17 - SWOT Analysis:**
- "SWOT_Analysis_Text" → Full SWOT analysis
- "Strengths_List" → List of strengths
- "Weaknesses_List" → List of weaknesses
- "Opportunities_List" → List of opportunities
- "Threats_List" → List of threats

**Section 18 - Risk Mitigation:**
- "Risk_Implementation_Text" → Risks during implementation
- "Risk_Operations_Text" → Risks during operations
- "Mitigation_Measures_Text" → Mitigation measures

**Section 21 - Conclusion:**
- "Conclusion_Text" → Final conclusion paragraph

═══════════════════════════════════════════════════════════════
USER INPUT PATTERNS FOR TEXT SECTIONS
═══════════════════════════════════════════════════════════════

**Keywords that indicate text sections:**
- "under [section] write/add/mention:"
- "description of [section]:"
- "[section] details:"
- "add paragraph about:"
- "include text:"
- "write about:"

**Examples:**

Input: "Under introduction write: This CFC will serve 50 textile units in Hyderabad cluster"
Output: "Introduction_Text" = "This CFC will serve 50 textile units in Hyderabad cluster"

Input: "Technology description: We will use automated looms with IoT sensors for real-time monitoring"
Output: "Technology_Text" = "We will use automated looms with IoT sensors for real-time monitoring"

Input: "Add SWOT strengths: strong market demand, skilled workforce, government support"
Output: "Strengths_List" = ["strong market demand", "skilled workforce", "government support"]

Input: "Conclusion should say: This project will transform the textile sector in Telangana"
Output: "Conclusion_Text" = "This project will transform the textile sector in Telangana"

Input: "Raw materials para: Cotton will be sourced from local farmers. Dyes imported from Germany."
Output: "Raw_materials_Text" = "Cotton will be sourced from local farmers. Dyes imported from Germany."

═══════════════════════════════════════════════════════════════
TABLE FIELDS (FROM PREVIOUS VERSION - KEEP ALL)
═══════════════════════════════════════════════════════════════

[All previous table field mappings remain the same - Total, Land_and_Building, etc.]

**Project Cost Table (Section 10i):**
- "Total" = "500"
- "Land_and_Building" = "100"
- "Plant_and_Machinery" = "200"
- "Preliminary_&_Pre-operative_expenses" = "25"
- "Margin_money_for_Working_Capital" = "50"

**Manpower Table (Section 8):**
- "Description_of_employee_1" = "Manager"
- "Number_of_employee_1" = "2"
- "Description_of_employee_2" = "Technician"
- "Number_of_employee_2" = "5"

**Machinery Table (Section 10iii):**
- "Machinery_Description_1" = "CNC Machine"
- "Machinery_No_1" = "2"
- "Machinery_Amount_1" = "50"

**Financial Viability (Section 14):**
- "Net_Block" = ["100", "95", "90", "85", "80"]
- "Income" = ["200", "250", "300", "350", "400"]
- "Profit_after_tax" = ["20", "30", "40", "50", "60"]

**Before/After Intervention (Section 15):**
- "Units_Before_Intervention" = "50"
- "Units_After_Intervention" = "100"
- "Employment_Before_Intervention" = "200"
- "Employment_After_Intervention" = "500"

═══════════════════════════════════════════════════════════════
EXTRACTION PRIORITY
═══════════════════════════════════════════════════════════════

1. **Extract ALL text sections first** (paragraphs, descriptions, lists)
2. **Then extract table data** (costs, manpower, machinery)
3. **Then extract basic info** (project name, candidate, address)

This ensures custom text is not missed!

═══════════════════════════════════════════════════════════════
OUTPUT FORMAT (ENHANCED)
═══════════════════════════════════════════════════════════════

Return ONLY valid JSON:
{
  "project_name": "Extracted project name",
  "candidate_name": "Extracted person/company name",
  "address": "Full address",
  "fields": {
    "Total": "500",
    "Land_and_Building": "100",
    "Introduction_Text": "This CFC will revolutionize textile manufacturing in Hyderabad...",
    "Technology_Text": "We will use Industry 4.0 automated systems...",
    "SWOT_Analysis_Text": "Strengths include strong market demand...",
    "Strengths_List": ["strong demand", "skilled workforce"],
    "Conclusion_Text": "This project will create 500 jobs and boost exports...",
    "Description_of_employee_1": "Manager",
    "Number_of_employee_1": "2"
  }
}

**CRITICAL RULES:**
1. Look for text section indicators: "write", "add", "mention", "description", "under", "about"
2. Map to correct _Text or _List field names
3. For lists, use JSON arrays
4. For paragraphs, keep as single strings (use \\n\\n for line breaks)
5. Don't lose any table data while adding text sections
6. Convert all numbers to strings
7. Be intelligent about section matching:
   - "introduction" → Introduction_Text
   - "tech details" → Technology_Text
   - "SWOT" → SWOT_Analysis_Text
   - "conclusion" → Conclusion_Text

Extract EVERYTHING from user input - both structured data AND custom text!"""

    try:
        response = client.chat.completions.create(
            model="openai/gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"User input:\n\n{user_input}\n\nExtract ALL information including custom text sections, lists, and table data. Return only JSON."}
            ],
            temperature=0.2,
            max_tokens=3000  # Increased for longer text sections
        )
        
        llm_output = response.choices[0].message.content.strip()
        
        # Extract JSON from response
        if "```json" in llm_output:
            json_str = llm_output.split("```json")[1].split("```")[0].strip()
        elif "```" in llm_output:
            json_str = llm_output.split("```")[1].split("```")[0].strip()
        else:
            start = llm_output.find('{')
            end = llm_output.rfind('}') + 1
            if start != -1 and end > start:
                json_str = llm_output[start:end]
            else:
                json_str = llm_output
        
        parsed_data = json.loads(json_str)
        return parsed_data
        
    except json.JSONDecodeError as e:
        print(f"❌ JSON Parse Error: {e}")
        print(f"LLM Output: {llm_output}")
        return None
    except Exception as e:
        print(f"❌ Error calling LLM: {e}")
        return None


def generate_pdf_from_natural_input(user_input):
    """
    Main function with enhanced display for text sections
    """
    
    print("\n" + "="*70)
    print("  🤖 INTELLIGENT DPR PDF GENERATOR (WITH CUSTOM TEXT SUPPORT)")
    print("="*70)
    print("\n📊 Analyzing your input with AI (GPT-4)...\n")
    
    # Get LLM mapping
    mapped_data = get_llm_mapping(user_input)
    
    if not mapped_data:
        print("❌ Failed to process input with LLM")
        return None
    
    # Extract basic info
    project_name = mapped_data.get("project_name", "Untitled Project")
    candidate_name = mapped_data.get("candidate_name", "Unknown")
    address = mapped_data.get("address", "Not specified")
    fields = mapped_data.get("fields", {})
    
    print("✅ AI successfully extracted the following information:")
    print(f"\n📋 Project Name: {project_name}")
    print(f"👤 Candidate/SPV: {candidate_name}")
    print(f"📍 Address: {address}")
    print(f"\n📊 Total fields extracted: {len(fields)}\n")
    
    # Categorize fields
    text_fields = {}
    list_fields = {}
    cost_fields = {}
    employee_fields = {}
    machinery_fields = {}
    financial_fields = {}
    other_fields = {}
    
    for key, value in fields.items():
        if isinstance(value, list) and not any(x in key for x in ['Net_Block', 'Income', 'Profit']):
            list_fields[key] = value
        elif '_Text' in key or '_Description' in key or 'Analysis' in key or key == 'Conclusion_Text':
            text_fields[key] = value
        elif any(x in key for x in ['Total', 'Land', 'Plant', 'Preliminary', 'Margin']):
            cost_fields[key] = value
        elif 'employee' in key:
            employee_fields[key] = value
        elif 'Machinery' in key:
            machinery_fields[key] = value
        elif any(x in key for x in ['Net_Block', 'Income', 'Profit', 'Current_Assets', 'Before_Intervention', 'After_Intervention']):
            financial_fields[key] = value
        else:
            other_fields[key] = value
    
    # Display text sections prominently
    if text_fields:
        print("  📝 CUSTOM TEXT SECTIONS ADDED:")
        for k, v in text_fields.items():
            section_name = k.replace('_Text', '').replace('_', ' ')
            preview = str(v)[:100] + "..." if len(str(v)) > 100 else str(v)
            print(f"     • {section_name}:")
            print(f"       '{preview}'")
    
    if list_fields:
        print("\n  📋 LISTS/BULLET POINTS:")
        for k, v in list_fields.items():
            section_name = k.replace('_List', '').replace('_', ' ')
            print(f"     • {section_name}:")
            for item in v:
                print(f"       - {item}")
    
    if cost_fields:
        print("\n  💰 PROJECT COST FIELDS:")
        for k, v in cost_fields.items():
            print(f"     • {k}: ₹{v} lakhs")
    
    if employee_fields:
        print("\n  👥 MANPOWER DETAILS:")
        emp_dict = {}
        for k, v in employee_fields.items():
            if 'Description' in k:
                num = k.split('_')[-1]
                emp_dict[num] = {'desc': v}
            elif 'Number' in k:
                num = k.split('_')[-1]
                if num not in emp_dict:
                    emp_dict[num] = {}
                emp_dict[num]['count'] = v
        for num in sorted(emp_dict.keys()):
            if 'desc' in emp_dict[num] and 'count' in emp_dict[num]:
                print(f"     • {emp_dict[num]['desc']}: {emp_dict[num]['count']} persons")
    
    if machinery_fields:
        print("\n  🔧 MACHINERY DETAILS:")
        mach_dict = {}
        for k, v in machinery_fields.items():
            num = k.split('_')[-1]
            if num not in mach_dict:
                mach_dict[num] = {}
            if 'Description' in k:
                mach_dict[num]['desc'] = v
            elif 'No' in k:
                mach_dict[num]['count'] = v
            elif 'Amount' in k:
                mach_dict[num]['amount'] = v
        for num in sorted(mach_dict.keys()):
            if 'desc' in mach_dict[num]:
                desc = mach_dict[num]['desc']
                count = mach_dict[num].get('count', '?')
                amount = mach_dict[num].get('amount', '?')
                print(f"     • {desc}: {count} units @ ₹{amount} lakhs")
    
    if financial_fields:
        print("\n  💹 FINANCIAL PROJECTIONS:")
        for k, v in financial_fields.items():
            if isinstance(v, list):
                print(f"     • {k}: {v}")
            else:
                print(f"     • {k}: {v}")
    
    if other_fields:
        print("\n  📝 OTHER FIELDS:")
        for k, v in other_fields.items():
            print(f"     • {k}: {v}")
    
    print("\n" + "="*70)
    print("🔄 Generating PDF with all extracted data...")
    print("="*70 + "\n")
    
    # Call the original PDF generation function
    try:
        pdf_path = create_dpr_pdf(
            project_name=project_name,
            candidate_name=candidate_name,
            address=address,
            **fields
        )
        
        print("\n" + "="*70)
        print("  ✅✅✅ SUCCESS! PDF GENERATED WITH CUSTOM TEXT")
        print("="*70)
        print(f"\n📄 PDF Location: {pdf_path}")
        print(f"📂 Folder: {os.path.dirname(pdf_path)}")
        
        if text_fields:
            print(f"\n✨ Your custom text was added to {len(text_fields)} sections!")
        if list_fields:
            print(f"📋 {len(list_fields)} lists/bullet points added!")
        
        print(f"\n💡 Total {len(fields)} fields mapped from your natural language input")
        print(f"\n🎉 Open the PDF to see your data with custom text sections!\n")
        
        return pdf_path
        
    except Exception as e:
        print(f"\n❌ Error generating PDF: {e}")
        import traceback
        traceback.print_exc()
        return None


def interactive_mode():
    """
    Interactive mode with enhanced examples including text sections
    """
    print("\n" + "="*70)
    print("  🚀 INTELLIGENT DPR GENERATOR - WITH CUSTOM TEXT SUPPORT")
    print("="*70)
    print("\n💬 Describe your project naturally - including custom text sections!\n")
    print("📝 NEW CAPABILITIES:\n")
    
    print("   ✨ Add Custom Text Under Any Heading:")
    print("      'Under introduction write: This CFC will serve textile cluster'")
    print("      'Technology description: Automated looms with IoT sensors'")
    print("      'Add conclusion: This project will create 500 jobs'\n")
    
    print("   📋 Add Bullet Points/Lists:")
    print("      'SWOT strengths: market demand, skilled workers, govt support'")
    print("      'Benefits include: higher efficiency, better quality, lower costs'\n")
    
    print("   📄 Complete Example:\n")
    print("   'Project: Textile Manufacturing CFC, Hyderabad")
    print("    SPV: Hyderabad Textiles Ltd, Address: Secunderabad, Telangana")
    print("    ")
    print("    Total project cost: 500 lakhs")
    print("    Land: 100 lakhs, Machinery: 200 lakhs")
    print("    Working capital: 50 lakhs, Preliminary: 25 lakhs")
    print("    ")
    print("    Under introduction write: This CFC will revolutionize the textile")
    print("    manufacturing sector in Hyderabad by providing state-of-art facilities")
    print("    to 50 MSME units in the cluster.")
    print("    ")
    print("    Technology description: We will deploy automated power looms with")
    print("    real-time IoT monitoring, Industry 4.0 compliant systems.")
    print("    ")
    print("    SWOT strengths: Strong market demand, skilled workforce available,")
    print("    government support confirmed")
    print("    ")
    print("    Manpower: 2 managers, 5 technicians, 10 workers")
    print("    Expected profit: 80 lakhs after intervention")
    print("    ")
    print("    Conclusion: This project will create 500 direct and indirect jobs,")
    print("    increase cluster turnover by 200%, and boost textile exports.'\n")
    
    print("="*70)
    print("\n📋 Enter your project details (type END on new line when done):\n")
    
    lines = []
    while True:
        try:
            line = input()
            if line.strip().upper() == "END":
                break
            lines.append(line)
        except EOFError:
            break
    
    user_input = "\n".join(lines)
    
    if not user_input.strip():
        print("\n❌ No input provided!")
        return
    
    # Generate PDF
    generate_pdf_from_natural_input(user_input)


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        if sys.argv[1] == "--test":
            # Enhanced test with custom text
            test_input = """
            Project name: Textile Manufacturing CFC
            My name is Kiran Kumar, Address: Hyderabad, Telangana
            
            Total project cost: 500 lakhs breakdown:
            - Land and building: 100 lakhs
            - Machinery: 200 lakhs
            - Preliminary expenses: 25 lakhs
            - Working capital: 50 lakhs
            
            Under introduction write: This Common Facility Centre will serve 50 textile MSME units 
            in the Hyderabad cluster. The project aims to provide world-class facilities including 
            automated looms, testing labs, and design studios.
            
            Technology description: We will deploy Industry 4.0 compliant automated power looms 
            with real-time IoT monitoring systems. The facility will have CAD/CAM design software 
            and modern testing equipment for quality control.
            
            SWOT strengths: Strong domestic and export market demand, availability of skilled workforce, 
            confirmed government support, proximity to raw material sources
            
            Manpower needed: 2 managers, 5 technicians, 3 supervisors, 10 workers
            
            Expected profit after intervention: 80 lakhs annually
            
            Add conclusion: This project will transform the textile manufacturing ecosystem in Telangana, 
            creating 500 direct and indirect employment opportunities, increasing cluster turnover by 200%, 
            and significantly boosting textile exports from the region.
            """
            print("🧪 Running enhanced test with custom text sections...\n")
            generate_pdf_from_natural_input(test_input)
        elif sys.argv[1] == "--help":
            print("\nUsage:")
            print("  python script.py          # Interactive mode")
            print("  python script.py --test   # Run test with sample data")
            print("  python script.py --help   # Show this help")
            print("\nFeatures:")
            print("  • Add custom text under any section")
            print("  • Create bullet point lists")
            print("  • Fill all table fields")
            print("  • Natural language input")
    else:
        # Interactive mode
        interactive_mode()