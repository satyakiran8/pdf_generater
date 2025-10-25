import os
import json
from openai import OpenAI

# Import the original PDF generation function
try:
    from dpr_pdf_generator import create_dpr_pdf
except ImportError:
    print("Error: Please save the PDF generator as 'dpr_pdf_generator.py'")
    exit(1)

# Initialize OpenAI client
client = OpenAI(
    base_url="https://openrouter.ai/api/v1",
    api_key="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"
)

def get_llm_mapping(user_input):
    """
    MEGA-ENHANCED LLM with support for:
    - Custom text under any heading/section
    - Bullet points and lists
    - Multi-paragraph descriptions
    - All existing table fields
    """
    
    system_prompt = """You are an ULTRA-INTELLIGENT DPR document analyzer that extracts BOTH structured table data AND custom text sections from natural language.

═══════════════════════════════════════════════════════════════
🆕 MASSIVE UPGRADE: CUSTOM TEXT ANYWHERE!
═══════════════════════════════════════════════════════════════

Users can now add custom text under ANY section heading using natural patterns:

**PATTERN 1: Direct Text Assignment**
Input: "Introduction text: This CFC will revolutionize textile manufacturing"
Output: "Introduction_Text" = "This CFC will revolutionize textile manufacturing"

**PATTERN 2: "Under X write/add/mention"**
Input: "Under technical aspects write: We will deploy Industry 4.0 systems"
Output: "Technical_Aspects_Text" = "We will deploy Industry 4.0 systems"

**PATTERN 3: "X description/details"**
Input: "Technology description: Automated looms with IoT sensors"
Output: "Technology_Text" = "Automated looms with IoT sensors"

**PATTERN 4: Bullet Points/Lists**
Input: "SWOT strengths: market demand, skilled workers, govt support"
Output: "Strengths_List" = ["market demand", "skilled workers", "govt support"]

**PATTERN 5: Multi-paragraph Text**
Input: "Add conclusion: Para 1 text here. Para 2 text here."
Output: "Conclusion_Text" = "Para 1 text here.\n\nPara 2 text here."

═══════════════════════════════════════════════════════════════
📝 COMPLETE FIELD MAPPING (All Sections)
═══════════════════════════════════════════════════════════════

**SECTION 2.1 - INTRODUCTION:**
- "Introduction_Text" → General introduction paragraph
- "Introduction_General_Text" → Under "General scenario of industrial growth"
- "Introduction_Sector_Text" → Under "Sector for which CFC is proposed"
- "Introduction_Cluster_Text" → Under "Cluster and its products"
- "Introduction_Relevance_Text" → Under "How CFC is relevant"

**SECTION 4 - PROMOTER DETAILS:**
- "Brief_bio_data_of_Promoters" → Promoter biography
- "Brief_about_Compliance_with_KYC_guidelines" → KYC compliance details
- "Management_Setup_Text" → Management structure details
- "Details_of_connected_lending" → Connected lending information
- "Adverse_auditors_remarks" → Auditor remarks
- "Particulars_of_previous_assistance" → Previous financial assistance
- "Pending_court_cases" → Court case details

**SECTION 6 - IMPLEMENTATION:**
- "Role_of_Implementing_Agency" → Agency responsibilities
- "Implementation_Period" → Timeline
- "Commitment_State_Govt" → State govt contribution

**SECTION 8 - TECHNICAL ASPECTS:**
- "Scope_of_project_Text" → Project scope description
- "Locational_details_Text" → Location information
- "Technology_Text" → Technology description
- "Industry_4_0_Text" → Industry 4.0 provisions
- "Raw_materials_Text" → Raw material details
- "Utilities_Text" → Utilities description
- "Power_Text" → Power requirements
- "Water_Text" → Water requirements
- "Effluent_disposal_Text" → Effluent handling
- "Manpower_Text" → Manpower details

**SECTION 10 - PROJECT COMPONENTS:**
- "Land_Building_Details_Text" → Land/building description
- "Misc_fixed_assets_Text" → Misc assets description
- "Preliminary_expenses_Text" → Preliminary expenses details
- "Pre_operative_expenses_Text" → Pre-operative expenses
- "Contingency_Provisions_Text" → Contingency details
- "Working_Capital_Text" → Working capital description
- "Usage_Charges_Text" → How facilities will be used

**SECTION 11 - FUND REQUIREMENT:**
- "Fund_requirement_analysis_Text" → Fund availability analysis

**SECTION 12 - USAGE CHARGES:**
- "Usage_Charges_Details" → Usage charge structure

**SECTION 13 - COMMERCIAL VIABILITY:**
- "Commercial_viability_comments" → Viability assessment

**SECTION 17 - SWOT ANALYSIS:**
- "SWOT_Analysis_Text" → Complete SWOT narrative
- "Strengths_List" → ["strength 1", "strength 2"]
- "Weaknesses_List" → ["weakness 1", "weakness 2"]
- "Opportunities_List" → ["opportunity 1", "opportunity 2"]
- "Threats_List" → ["threat 1", "threat 2"]

**SECTION 18 - RISK MITIGATION:**
- "Risk_Implementation_Text" → Implementation risks
- "Risk_Operations_Text" → Operational risks
- "Mitigation_Measures_Text" → How to mitigate risks

**SECTION 19 - ECONOMICS:**
- "DSCR_Analysis_Text" → Debt service coverage ratio details
- "Balance_Sheet_Text" → Balance sheet notes
- "Break_Even_Analysis_Text" → Break-even explanation

**SECTION 20 - COMMERCIAL VIABILITY:**
- "ROCE_Text" → Return on capital employed
- "Sensitivity_Analysis_Text" → Sensitivity analysis

**SECTION 21 - CONCLUSION:**
- "Conclusion_Text" → Final conclusion paragraph

═══════════════════════════════════════════════════════════════
🎯 SMART KEYWORD DETECTION
═══════════════════════════════════════════════════════════════

**Look for these indicators of custom text:**
- "write", "add", "mention", "include", "describe"
- "under [section]:", "for [section]:", "[section] details:"
- "description:", "paragraph:", "text:"
- "strengths:", "weaknesses:", "benefits:", "objectives:"
- "conclusion:", "summary:", "overview:"
- **CRITICAL: "under 17 section", "add under section 17", "17 section heading"**

**Section Name Normalization:**
- "introduction" → Introduction_Text
- "tech details" → Technology_Text  
- "technology" → Technology_Text
- "SWOT" → SWOT_Analysis_Text
- **"17 section" → SWOT_Analysis_Text**
- **"section 17" → SWOT_Analysis_Text**
- **"under 17" → SWOT_Analysis_Text**
- "conclusion" → Conclusion_Text
- "raw materials" → Raw_materials_Text
- "implementation" → Role_of_Implementing_Agency

═══════════════════════════════════════════════════════════════
📊 TABLE FIELDS (Keep All Existing)
═══════════════════════════════════════════════════════════════

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
🔍 EXTRACTION LOGIC (PRIORITY ORDER)
═══════════════════════════════════════════════════════════════

**STEP 1: Extract ALL custom text sections FIRST**
Look for: "under X write Y", "X description: Y", "add text about X"
Output format: Section_Name_Text or Section_Name_List

**STEP 2: Extract table data**
Cost breakdowns, manpower, machinery, financial projections

**STEP 3: Extract basic info**
Project name, candidate name, address

═══════════════════════════════════════════════════════════════
💡 EXAMPLES OF TEXT EXTRACTION
═══════════════════════════════════════════════════════════════

**Example 1: Simple paragraph**
Input: "Under introduction write: This CFC will serve 50 MSME units in Hyderabad textile cluster"
Output: {
  "fields": {
    "Introduction_Text": "This CFC will serve 50 MSME units in Hyderabad textile cluster"
  }
}

**Example 2: KYC Compliance**
Input: "Brief about Compliance with KYC guidelines=All KYC documents verified and approved"
Output: {
  "fields": {
    "Brief_about_Compliance_with_KYC_guidelines": "All KYC documents verified and approved"
  }
}

**Example 3: Technology description**
Input: "Technology description: We will deploy automated power looms with Industry 4.0 IoT sensors"
Output: {
  "fields": {
    "Technology_Text": "We will deploy automated power looms with Industry 4.0 IoT sensors"
  }
}

**Example 3: SWOT with lists**
Input: "SWOT strengths: Strong market demand, skilled workforce, government support. Weaknesses: Limited capital, old infrastructure"
Output: {
  "fields": {
    "Strengths_List": ["Strong market demand", "skilled workforce", "government support"],
    "Weaknesses_List": ["Limited capital", "old infrastructure"]
  }
}

**Example 4: Section 4 fields**
Input: "Brief about Compliance with KYC guidelines=All documents verified. Details of connected lending=None found. Management Setup Text=CEO and 3 directors"
Output: {
  "fields": {
    "Brief_about_Compliance_with_KYC_guidelines": "All documents verified",
    "Details_of_connected_lending": "None found",
    "Management_Setup_Text": "CEO and 3 directors"
  }
}

**Example 5: Section 17 (SWOT) - CRITICAL**
Input: "add under 17 section heading=this is not mandatory"
Output: {
  "fields": {
    "SWOT_Analysis_Text": "this is not mandatory"
  }
}

**Example 6: Section 17 variations**
Input: "add under section 17=SWOT analysis is comprehensive"
Output: {
  "fields": {
    "SWOT_Analysis_Text": "SWOT analysis is comprehensive"
  }
}

Input: "17 section heading text=Market analysis shows strong potential"
Output: {
  "fields": {
    "SWOT_Analysis_Text": "Market analysis shows strong potential"
  }
}

**Example 7: Conclusion paragraph**
Input: "Add conclusion: This project will create 500 jobs and boost textile exports by 200%"
Output: {
  "fields": {
    "Conclusion_Text": "This project will create 500 jobs and boost textile exports by 200%"
  }
}

**Example 5: Multi-section with tables**
Input: "Project cost 500 lakhs. Land 100, machinery 200. Under introduction: This is a textile CFC. Technology: Automated looms with IoT. Brief about Compliance with KYC guidelines=Verified by bank"
Output: {
  "fields": {
    "Total": "500",
    "Land_and_Building": "100",
    "Plant_and_Machinery": "200",
    "Introduction_Text": "This is a textile CFC",
    "Technology_Text": "Automated looms with IoT",
    "Brief_about_Compliance_with_KYC_guidelines": "Verified by bank"
  }
}

═══════════════════════════════════════════════════════════════
⚙️ OUTPUT FORMAT
═══════════════════════════════════════════════════════════════

Return ONLY valid JSON:
{
  "project_name": "Extracted name",
  "candidate_name": "Person/SPV name",
  "address": "Full address",
  "fields": {
    "Total": "500",
    "Introduction_Text": "Custom intro paragraph...",
    "Technology_Text": "Tech description...",
    "SWOT_Analysis_Text": "Full SWOT...",
    "Strengths_List": ["strength1", "strength2"],
    "Conclusion_Text": "Final conclusion...",
    "Description_of_employee_1": "Manager",
    "Number_of_employee_1": "2"
  }
}

**CRITICAL RULES:**
1. **Text sections take priority** - Extract these FIRST
2. **Use exact field names** - Introduction_Text, Technology_Text, Conclusion_Text
3. **For lists** - Use JSON arrays: ["item1", "item2"]
4. **For paragraphs** - Keep as single string with \\n\\n for breaks
5. **All numbers to strings** - "500" not 500
6. **Be intelligent about matching:**
   - "intro text" → Introduction_Text
   - "tech info" → Technology_Text
   - "strengths" → Strengths_List
   - "conclusion para" → Conclusion_Text
   - "Brief about Compliance with KYC guidelines" → Brief_about_Compliance_with_KYC_guidelines
   - "KYC compliance" → Brief_about_Compliance_with_KYC_guidelines
   - "connected lending" → Details_of_connected_lending
7. **Extract EVERYTHING** - Don't skip any data!
8. **CRITICAL SECTION DETECTION:**
   - If input contains "17 section" OR "section 17" OR "under 17" → ALWAYS map to SWOT_Analysis_Text
   - If input contains "21 section" OR "section 21" → ALWAYS map to Conclusion_Text
   - If input contains "8 section" OR "section 8" → Check subsection (technology, raw materials, etc.)
   - If input contains "4 section" OR "section 4" → Check subsection (KYC, promoters, etc.)

**Pattern Matching Priority:**
1. **Exact section number match** (highest priority)
   - "add under 17 section heading" → SWOT_Analysis_Text
   - "section 17 =" → SWOT_Analysis_Text
2. **Section keyword match**
   - "SWOT" → SWOT_Analysis_Text
   - "conclusion" → Conclusion_Text
3. **Descriptive match**
   - "technology description" → Technology_Text

═══════════════════════════════════════════════════════════════

Now extract ALL information from user input including custom text sections!"""

    try:
        response = client.chat.completions.create(
            model="openai/gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": f"""Extract ALL information from this input. Pay special attention to section numbers!

USER INPUT:
{user_input}

RULES:
- If you see "17 section" or "section 17" or "under 17" → Use field "SWOT_Analysis_Text"
- If you see "21 section" or "section 21" → Use field "Conclusion_Text"
- If you see "4 section" or "section 4" with KYC → Use field "Brief_about_Compliance_with_KYC_guidelines"
- Always extract the text after "=" sign as the value

Return ONLY valid JSON with both table fields AND text sections."""}
            ],
            temperature=0.1,  # Lower temperature for more consistent extraction
            max_tokens=4000
        )
        
        llm_output = response.choices[0].message.content.strip()
        
        # Extract JSON
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
    Main function with MEGA-ENHANCED display for text sections
    """
    
    print("\n" + "="*80)
    print("  🚀 MEGA-INTELLIGENT DPR GENERATOR - TEXT SECTIONS SUPPORT")
    print("="*80)
    print("\n🤖 Analyzing your input with GPT-4 (Enhanced for custom text)...\n")
    
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
    
    print("✅ AI successfully extracted information:")
    print(f"\n📋 Project Name: {project_name}")
    print(f"👤 Candidate/SPV: {candidate_name}")
    print(f"📍 Address: {address}")
    print(f"\n📊 Total fields extracted: {len(fields)}\n")
    
    # Categorize fields with priority to text sections
    text_sections = {}
    list_sections = {}
    cost_fields = {}
    employee_fields = {}
    machinery_fields = {}
    financial_fields = {}
    other_fields = {}
    
    for key, value in fields.items():
        # Text sections (highest priority display)
        if '_Text' in key or 'Analysis' in key or 'Description' in key or key == 'Conclusion_Text':
            text_sections[key] = value
        # Lists
        elif '_List' in key or (isinstance(value, list) and not any(x in key for x in ['Net_Block', 'Income', 'Profit'])):
            list_sections[key] = value
        # Tables
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
    
    # Display with PRIORITY for text sections
    if text_sections:
        print("  ✨ CUSTOM TEXT SECTIONS (NEW!):")
        print("  " + "="*76)
        for k, v in text_sections.items():
            section_name = k.replace('_Text', '').replace('_', ' ').title()
            # Show first 150 chars
            preview = str(v)[:150] + "..." if len(str(v)) > 150 else str(v)
            print(f"\n  📝 {section_name}:")
            print(f"     \"{preview}\"")
        print("\n  " + "="*76)
    
    if list_sections:
        print("\n  📋 LISTS & BULLET POINTS:")
        for k, v in list_sections.items():
            section_name = k.replace('_List', '').replace('_', ' ').title()
            print(f"\n     • {section_name}:")
            for item in v:
                print(f"       - {item}")
    
    if cost_fields:
        print("\n  💰 PROJECT COST BREAKDOWN:")
        for k, v in cost_fields.items():
            print(f"     • {k.replace('_', ' ')}: ₹{v} lakhs")
    
    if employee_fields:
        print("\n  👥 MANPOWER REQUIREMENTS:")
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
        print("\n  🔧 MACHINERY SPECIFICATIONS:")
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
                print(f"     • {k.replace('_', ' ')}: {v}")
            else:
                print(f"     • {k.replace('_', ' ')}: {v}")
    
    if other_fields:
        print("\n  📝 OTHER FIELDS:")
        for k, v in other_fields.items():
            preview = str(v)[:80] + "..." if len(str(v)) > 80 else str(v)
            print(f"     • {k.replace('_', ' ')}: {preview}")
    
    print("\n" + "="*80)
    print("🔄 Generating PDF with ALL extracted data (including custom text)...")
    print("="*80 + "\n")
    
    # Call PDF generation
    try:
        pdf_path = create_dpr_pdf(
            project_name=project_name,
            candidate_name=candidate_name,
            address=address,
            **fields
        )
        
        print("\n" + "="*80)
        print("  ✅✅✅ SUCCESS! PDF GENERATED WITH CUSTOM TEXT SECTIONS")
        print("="*80)
        print(f"\n📄 PDF Location: {pdf_path}")
        print(f"📂 Folder: {os.path.dirname(pdf_path)}")
        
        if text_sections:
            print(f"\n✨ Custom text added to {len(text_sections)} sections!")
        if list_sections:
            print(f"📋 {len(list_sections)} lists with bullet points added!")
        if cost_fields:
            print(f"💰 {len(cost_fields)} cost fields filled!")
        
        print(f"\n💡 Total {len(fields)} fields successfully mapped!")
        print(f"\n🎉 Open the PDF to see your complete DPR with custom sections!\n")
        
        return pdf_path
        
    except Exception as e:
        print(f"\n❌ Error generating PDF: {e}")
        import traceback
        traceback.print_exc()
        return None


def interactive_mode():
    """
    Interactive mode with MEGA-ENHANCED examples
    """
    print("\n" + "="*80)
    print("  🚀 MEGA-INTELLIGENT DPR GENERATOR")
    print("  ✨ NOW WITH CUSTOM TEXT SECTIONS SUPPORT!")
    print("="*80)
    print("\n💬 Describe your project naturally - INCLUDING custom text!\n")
    
    print("🎯 NEW CAPABILITIES:\n")
    
    print("   ✨ Add Custom Text Under ANY Heading:")
    print("      'Under introduction write: This CFC will transform textile sector'")
    print("      'Technology description: Automated looms with IoT monitoring'")
    print("      'Add conclusion: Project will create 500 jobs'\n")
    
    print("   📋 Add Bullet Points/Lists:")
    print("      'SWOT strengths: market demand, skilled workers, govt support'")
    print("      'Benefits: higher efficiency, better quality, lower costs'\n")
    
    print("   📊 Fill Tables (as before):")
    print("      'Total cost 500 lakhs, Land 100, Machinery 200'")
    print("      'Manpower: 2 managers, 5 technicians'\n")
    
    print("="*80)
    print("\n📝 COMPLETE EXAMPLE WITH EVERYTHING:\n")
    
    example = """
Project: Textile Manufacturing CFC, Hyderabad
SPV: Hyderabad Textiles Pvt Ltd
Address: Industrial Area Phase 2, Secunderabad, Telangana

Total project cost: 500 lakhs breakdown:
- Land and building: 100 lakhs
- Plant and machinery: 200 lakhs
- Preliminary expenses: 25 lakhs
- Working capital: 50 lakhs

Under introduction write: This Common Facility Centre will revolutionize 
the textile manufacturing sector in Hyderabad. The project aims to serve 
50 MSME units in the cluster by providing world-class facilities including 
automated looms, testing laboratories, and design studios. This initiative 
will significantly boost the competitiveness of local manufacturers.

Technology description: We will deploy Industry 4.0 compliant automated 
power looms with real-time IoT monitoring systems. The facility will feature 
CAD/CAM design software, modern quality testing equipment, and cloud-based 
production management systems. All machinery will be energy-efficient and 
environmentally sustainable.

SWOT strengths: Strong domestic and export market demand, availability of 
skilled workforce in the region, confirmed government support under MSME 
schemes, proximity to raw material suppliers, existing cluster infrastructure

SWOT weaknesses: Limited initial capital, aging infrastructure in some units, 
need for technical training, competition from imported textiles

Manpower requirements: 2 managers, 5 senior technicians, 3 supervisors, 
10 machine operators

Expected profit after intervention: 80 lakhs per year

Add conclusion: This project represents a transformative opportunity for 
the Telangana textile sector. Upon completion, it will create 500 direct 
and indirect employment opportunities, increase cluster turnover by 200%, 
and significantly boost textile exports from the region. The project aligns 
perfectly with the government's Make in India initiative.
"""
    
    print(example)
    print("\n" + "="*80)
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
            test_input = """
            Project: Textile Manufacturing CFC
            SPV: Hyderabad Textiles Ltd
            Address: Secunderabad, Telangana
            
            Total cost: 500 lakhs
            Land: 100, Machinery: 200, Preliminary: 25, Working capital: 50
            
            Under introduction write: This CFC will serve 50 textile MSME units 
            in Hyderabad cluster with world-class automated facilities.
            
            Technology description: Industry 4.0 automated power looms with 
            real-time IoT monitoring and cloud-based production management.
            
            SWOT strengths: Strong market demand, skilled workforce, government support
            
            Manpower: 2 managers, 5 technicians, 10 workers
            
            Add conclusion: This project will create 500 jobs, increase turnover 
            by 200%, and boost textile exports significantly.
            """
            print("🧪 Running MEGA test with custom text...\n")
            generate_pdf_from_natural_input(test_input)
        elif sys.argv[1] == "--help":
            print("\n" + "="*80)
            print("  MEGA-INTELLIGENT DPR GENERATOR - HELP")
            print("="*80)
            print("\nUsage:")
            print("  python script.py          # Interactive mode")
            print("  python script.py --test   # Run test with sample data")
            print("  python script.py --help   # Show this help\n")
            print("Features:")
            print("  ✅ Add custom text under ANY section heading")
            print("  ✅ Create bullet point lists")
            print("  ✅ Fill all table fields")
            print("  ✅ Natural language input")
            print("  ✅ Intelligent AI extraction\n")
    else:
        interactive_mode()