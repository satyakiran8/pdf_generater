import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib import colors
from datetime import datetime

def get_optional_input(prompt, default="", allow_skip_all=False):
    """Helper function to get optional input"""
    if allow_skip_all:
        value = input(f"{prompt} (Press Enter to skip, type 'done' to finish and generate PDF): ").strip()
        if value.lower() == 'done':
            return 'DONE'
    else:
        value = input(f"{prompt} (Press Enter to skip): ").strip()
    return value if value else default

def get_section2_data():
    """Get Section 2 table data - constant data, no input needed"""
    return [
        ['Name of applicant,\ncontact details, etc', 'As CFC Registered address / administrative address may be different\nfrom CFC facilities address, the same may be provided'],
        ['Location of Common\nFacility Centre', 'Address where facilities are proposed may be provided'],
        ['Main facilities being\nproposed', 'Details of facilities to be provided']
    ]

def get_spv_data(candidate_name, address):
    """Get SPV Information table data"""
    print("\n=== Section 3: Information about SPV ===")
    print("(Type 'done' at any point to skip remaining fields and generate PDF)")
    name_address = f"{candidate_name}, {address}"
    
    spv_data = [
        ['S. No.', 'Description', 'Details/ Compliance'],
        ['(i)', 'Name and address', name_address],
    ]
    
    fields = [
        ('(ii)', 'Registration details of SPV\n(including registration as Section\n8 company under the Companies\nAct 2013)'),
        ('(iii)', 'Names of the State Govt and\nMSME officials in SPV'),
        ('(iv)', 'Date of formation of the\ncompany'),
        ('(v)', 'Date of commencement of\nbusiness'),
        ('(vi)', 'Number of MSE Member Units'),
        ('(vii)', 'Bye laws or MoA and AoA\nsubmitted'),
        ('(viii)', 'Main objectives of the SPV'),
        ('(ix)', 'SPV to have a character of\ninclusiveness wherein provision\nfor enrolling new members to\nenable prospective entrepreneurs\nin the cluster to utilise the\nfacility'),
        ('(x)', "Clause about 'Profits/surplus\n to be ploughed back\n to CFC' included or not"),
        ('(xi)', 'Authorized share capital'),
        ('(xii)', 'Shareholding Pattern\n(Annexure-3 to be filled in)'),
    ]
    
    for sno, desc in fields:
        value = get_optional_input(f"{sno} {desc.replace(chr(10), ' ')}", allow_skip_all=True)
        if value == 'DONE':
            spv_data.append([sno, desc, ''])
            # Fill remaining fields with empty strings
            remaining_idx = fields.index((sno, desc)) + 1
            for rem_sno, rem_desc in fields[remaining_idx:]:
                spv_data.append([rem_sno, rem_desc, ''])
            return spv_data, True
        spv_data.append([sno, desc, value])
    
    return spv_data, False

def get_spv_data2():
    """Get continuation of SPV table"""
    print("\n=== Section 3 Continued ===")
    spv_data2 = []
    
    fields = [
        ('(xiii)', 'Commitment letter for SPV\nUpfront contribution'),
        ('(xiv)', 'Project specific A/c in schedule\nA bank'),
        ('(xv)', "Clause about 'CFC may be\n utilised by SPV members as also\nothers in a cluster and \nEvidence for SPV members'\nability to utilise at least 60% of installed capacity'"),
        ('(xvi)', 'Main Role of SPV'),
        ('(xvii)', 'Trust building of SPV so that\nCFC may be successful'),
    ]
    
    for sno, desc in fields:
        value = get_optional_input(f"{sno} {desc.replace(chr(10), ' ')}", allow_skip_all=True)
        if value == 'DONE':
            spv_data2.append([sno, desc, ''])
            # Fill remaining fields with empty strings
            remaining_idx = fields.index((sno, desc)) + 1
            for rem_sno, rem_desc in fields[remaining_idx:]:
                spv_data2.append([rem_sno, rem_desc, ''])
            return spv_data2, True
        spv_data2.append([sno, desc, value])
    
    return spv_data2, False

def get_promoters_data():
    """Get promoters table data"""
    print("\n=== Section 4: Details of Project Promoters ===")
    num_promoters = input("How many promoters to add? (Press Enter for 0, type 'done' to skip): ").strip()
    
    if num_promoters.lower() == 'done':
        # Return empty promoters table
        return [
            ['Name of the\nOffice bearers\n of the SPV', '', '', '', '', '', '', ''],
            ['Age (years)', '', '', '', '', '', '', ''],
            ['Educational\n Qualification', '', '', '', '', '', '', ''],
            ['Relationship with \nthe chief promoter', '', '', '', '', '', '', ''],
            ['Experience in what\n capacity/ industry/ years', '', '', '', '', '', '', ''],
            ['Income Tax / Wealth \nTax Status \n(returns for 3 years\n to be furnished)', '', '', '', '', '', '', ''],
            ['Other concerns \ninterest / in which capacity \n/financial stake', '', '', '', '', '', '', '']
        ], True
    
    num_promoters = int(num_promoters) if num_promoters.isdigit() else 0
    
    promoters_data = [
        ['Name of the\nOffice bearers\n of the SPV', '', '', '', '', '', '', ''],
        ['Age (years)', '', '', '', '', '', '', ''],
        ['Educational\n Qualification', '', '', '', '', '', '', ''],
        ['Relationship with \nthe chief promoter', '', '', '', '', '', '', ''],
        ['Experience in what\n capacity/ industry/ years', '', '', '', '', '', '', ''],
        ['Income Tax / Wealth \nTax Status \n(returns for 3 years\n to be furnished)', '', '', '', '', '', '', ''],
        ['Other concerns \ninterest / in which capacity \n/financial stake', '', '', '', '', '', '', '']
    ]
    
    for i in range(min(num_promoters, 7)):  # Max 7 columns after first
        print(f"\n--- Promoter {i+1} ---")
        
        name = get_optional_input("Name", allow_skip_all=True)
        if name == 'DONE':
            return promoters_data, True
        promoters_data[0][i+1] = name
        
        age = get_optional_input("Age", allow_skip_all=True)
        if age == 'DONE':
            return promoters_data, True
        promoters_data[1][i+1] = age
        
        edu = get_optional_input("Educational Qualification", allow_skip_all=True)
        if edu == 'DONE':
            return promoters_data, True
        promoters_data[2][i+1] = edu
        
        rel = get_optional_input("Relationship with chief promoter", allow_skip_all=True)
        if rel == 'DONE':
            return promoters_data, True
        promoters_data[3][i+1] = rel
        
        exp = get_optional_input("Experience", allow_skip_all=True)
        if exp == 'DONE':
            return promoters_data, True
        promoters_data[4][i+1] = exp
        
        tax = get_optional_input("Tax Status", allow_skip_all=True)
        if tax == 'DONE':
            return promoters_data, True
        promoters_data[5][i+1] = tax
        
        concerns = get_optional_input("Other concerns", allow_skip_all=True)
        if concerns == 'DONE':
            return promoters_data, True
        promoters_data[6][i+1] = concerns
    
    return promoters_data, False

def get_implementation_data():
    """Get implementing arrangements data"""
    print("\n=== Section 6: Implementing Arrangements ===")
    
    impl_data = [
        ['Description', 'Compliance'],
        ['a. Name of Implementation Agency', get_optional_input("Implementation Agency Name")],
        ['b. Role of Implementing Agency (e.g. implementation and monitoring of project,\nsending regular progress reports, issuing proper UCs, )', get_optional_input("Role of Implementation Agency")],
        ['c. Implementation Period', get_optional_input("Implementation Period")],
        ['d. Commitment of State Government upfront contribution', get_optional_input("State Govt Commitment")],
        ['e. Commitment of Loans (Working capital and/ or term loan)', get_optional_input("Loan Commitment")],
    ]
    
    return impl_data

def get_manpower_data():
    """Get manpower table data"""
    print("\n=== Section 8: Manpower Details ===")
    num_employees = input("How many employee types to add? (Press Enter for 0): ").strip()
    num_employees = int(num_employees) if num_employees.isdigit() else 0
    
    manpower_data = [
        ['S. No.', 'Description of the employee', 'Number'],
    ]
    
    for i in range(max(num_employees, 4)):  # Minimum 4 rows
        if i < num_employees:
            desc = get_optional_input(f"Employee {i+1} Description")
            num = get_optional_input(f"Employee {i+1} Number")
            manpower_data.append([str(i+1), desc, num])
        else:
            manpower_data.append([str(i+1), '', ''])
    
    return manpower_data

def get_schedule_data():
    """Get implementation schedule data"""
    print("\n=== Section 9: Implementation Schedule ===")
    print("(Press Enter to skip any dates)")
    
    activities = [
        'Preparation of Project Report',
        'Sanction of Grant from Government of India',
        'NOC from Pollution Control Board',
        'Site Development',
        'Building up-keep',
        'Placement of order to equipment supplier',
        'Supply of equipments by suppliers',
        'Installation of equipments at site',
        'Sanction of power connection',
        'Trial Run',
        'Commercial Production',
    ]
    
    schedule_data = [['Activities', 'Start Date', 'Completion Date']]
    
    for activity in activities:
        start = get_optional_input(f"{activity} - Start Date")
        end = get_optional_input(f"{activity} - Completion Date")
        schedule_data.append([activity, start, end])
    
    return schedule_data

def get_cost_data():
    """Get project cost data"""
    print("\n=== Section 10: Estimated Project Cost ===")
    
    cost_data = [
        ['S. No.', 'Particulars', 'Amount'],
        ['1', 'Land and Building', get_optional_input("Land and Building (Rs. in lakh)")],
        ['2', 'Plant & Machinery including MFA,\nInstallation, Taxes/duties,\nContingencies, etc.', get_optional_input("Plant & Machinery (Rs. in lakh)")],
        ['3', 'Preliminary & Pre-operative expenses', get_optional_input("Preliminary expenses (Rs. in lakh)")],
        ['4', 'Margin money for Working Capital', get_optional_input("Working Capital (Rs. in lakh)")],
        ['', 'Total', get_optional_input("Total Project Cost (Rs. in lakh)")],
    ]
    
    return cost_data

def get_machinery_data():
    """Get machinery details"""
    print("\n=== Section 10: Plant & Machinery ===")
    num_items = input("How many machinery items to add? (Press Enter for 0): ").strip()
    num_items = int(num_items) if num_items.isdigit() else 0
    
    machinery_data = [['S. No.', 'Description', 'No.', 'Amount']]
    
    for i in range(max(num_items, 4)):  # Minimum 4 rows
        if i < num_items:
            desc = get_optional_input(f"Item {i+1} Description")
            qty = get_optional_input(f"Item {i+1} Quantity")
            amt = get_optional_input(f"Item {i+1} Amount (Rs. lakh)")
            machinery_data.append([str(i+1), desc, qty, amt])
        else:
            machinery_data.append([str(i+1), '', '', ''])
    
    return machinery_data

def get_financing_data():
    """Get means of financing data"""
    print("\n=== Section 10: Proposed Means of Financing ===")
    num_sources = input("How many financing sources to add? (Press Enter for 0): ").strip()
    num_sources = int(num_sources) if num_sources.isdigit() else 0
    
    financing_data = [['S. No.', 'Particulars', 'Percentage', 'Amount']]
    
    for i in range(num_sources):
        particular = get_optional_input(f"Source {i+1} Name")
        percentage = get_optional_input(f"Source {i+1} Percentage")
        amount = get_optional_input(f"Source {i+1} Amount")
        financing_data.append([str(i+1), particular, percentage, amount])
    
    total_amt = get_optional_input("Total Amount")
    financing_data.append(['', 'Total', '', total_amt])
    
    return financing_data

def get_financial_data():
    """Get financial viability data"""
    print("\n=== Section 14: Financial Economic Viability ===")
    print("Enter data for 5 financial years (Press Enter to skip)")
    
    params = [
        'Net Block',
        'Current Assets (incl. cash/bank balance)',
        'Current Liabilities (incl. principal installment falling due during the year)',
        'Long term borrowings',
        'Capital',
        'Reserves and Surplus',
        'Unsecured loan',
        'Net Worth (incl. GoI Subsidy as Quasi-equity)',
        'Income',
        'Gross profit',
        'Depreciation',
        'Profit after tax',
        'Gross Cash Accruals',
    ]
    
    financial_data = [['S. No.', 'Particulars', 'FY 1', 'FY 2', 'FY3', 'FY4', 'FY5']]
    
    for i, param in enumerate(params, 1):
        row = [str(i), param]
        for fy in range(1, 6):
            value = get_optional_input(f"{param} - FY{fy}")
            row.append(value)
        financial_data.append(row)
    
    return financial_data

def get_performance_data():
    """Get projected performance data"""
    print("\n=== Section 15: Projected Performance ===")
    
    particulars = [
        'Units (including details of SC/ST/Women/Minorities)',
        'Employment',
        'Production',
        'Exports',
        'Import Substitution',
        'Number of patent expected aimed',
        'Investment',
        'Turnover',
        'Profit',
        'Quality Certification',
        'Any others (No. of ZED certified units)',
    ]
    
    performance_data = [['Particulars', 'Before Intervention\nQty. / Outcome', 'After Intervention\nQty. / Outcome']]
    
    for particular in particulars:
        before = get_optional_input(f"{particular} - Before")
        after = get_optional_input(f"{particular} - After")
        performance_data.append([particular, before, after])
    
    return performance_data

def create_dpr_pdf(project_name, candidate_name, address):
    """
    DPR Template PDF ni create chestundi - interactive ga data fill cheskune option tho
    """
    
    # Get all interactive data
    section2_data = get_section2_data()
    
    spv_data, should_stop = get_spv_data(candidate_name, address)
    if should_stop:
        print("\n✓ Stopping data collection. Generating PDF with filled data...")
        spv_data2 = [
            ['(xiii)', 'Commitment letter for SPV\nUpfront contribution', ''],
            ['(xiv)', 'Project specific A/c in schedule\nA bank', ''],
            ['(xv)', "Clause about 'CFC may be\n utilised by SPV members as also\nothers in a cluster and \nEvidence for SPV members'\nability to utilise at least 60% of installed capacity'", ''],
            ['(xvi)', 'Main Role of SPV', ''],
            ['(xvii)', 'Trust building of SPV so that\nCFC may be successful', ''],
        ]
    else:
        spv_data2, should_stop = get_spv_data2()
        if should_stop:
            print("\n✓ Stopping data collection. Generating PDF with filled data...")
    
    if not should_stop:
        promoters_data, should_stop = get_promoters_data()
        if should_stop:
            print("\n✓ Stopping data collection. Generating PDF with filled data...")
    else:
        promoters_data = [
            ['Name of the\nOffice bearers\n of the SPV', '', '', '', '', '', '', ''],
            ['Age (years)', '', '', '', '', '', '', ''],
            ['Educational\n Qualification', '', '', '', '', '', '', ''],
            ['Relationship with \nthe chief promoter', '', '', '', '', '', '', ''],
            ['Experience in what\n capacity/ industry/ years', '', '', '', '', '', '', ''],
            ['Income Tax / Wealth \nTax Status \n(returns for 3 years\n to be furnished)', '', '', '', '', '', '', ''],
            ['Other concerns \ninterest / in which capacity \n/financial stake', '', '', '', '', '', '', '']
        ]
    
    # Set default empty data for remaining sections if stopped early
    if should_stop:
        impl_data = [
            ['Description', 'Compliance'],
            ['a. Name of Implementation Agency', ''],
            ['b. Role of Implementing Agency (e.g. implementation and monitoring of project,\nsending regular progress reports, issuing proper UCs, )', ''],
            ['c. Implementation Period', ''],
            ['d. Commitment of State Government upfront contribution', ''],
            ['e. Commitment of Loans (Working capital and/ or term loan)', ''],
        ]
        manpower_data = [
            ['S. No.', 'Description of the employee', 'Number'],
            ['1', '', ''],
            ['2', '', ''],
            ['3', '', ''],
            ['4', '', ''],
        ]
        schedule_data = [
            ['Activities', 'Start Date', 'Completion Date'],
            ['Preparation of Project Report', '', ''],
            ['Sanction of Grant from Government of India', '', ''],
            ['NOC from Pollution Control Board', '', ''],
            ['Site Development', '', ''],
            ['Building up-keep', '', ''],
            ['Placement of order to equipment supplier', '', ''],
            ['Supply of equipments by suppliers', '', ''],
            ['Installation of equipments at site', '', ''],
            ['Sanction of power connection', '', ''],
            ['Trial Run', '', ''],
            ['Commercial Production', '', ''],
        ]
        cost_data = [
            ['S. No.', 'Particulars', 'Amount'],
            ['1', 'Land and Building', ''],
            ['2', 'Plant & Machinery including MFA,\nInstallation, Taxes/duties,\nContingencies, etc.', ''],
            ['3', 'Preliminary & Pre-operative expenses', ''],
            ['4', 'Margin money for Working Capital', ''],
            ['', 'Total', ''],
        ]
        machinery_data = [
            ['S. No.', 'Description', 'No.', 'Amount'],
            ['1', '', '', ''],
            ['2', '', '', ''],
            ['3', '', '', ''],
            ['4', '', '', ''],
        ]
        financing_data = [
            ['S. No.', 'Particulars', 'Percentage', 'Amount'],
            ['', '', '', ''],
            ['', 'Total', '', ''],
        ]
        financial_data = [
            ['S. No.', 'Particulars', 'FY 1', 'FY 2', 'FY3', 'FY4', 'FY5'],
            ['1', 'Net Block', '', '', '', '', ''],
            ['2', 'Current Assets (incl. cash/bank balance)', '', '', '', '', ''],
            ['3', 'Current Liabilities (incl. principal installment falling due during the year)', '', '', '', '', ''],
            ['4', 'Long term borrowings', '', '', '', '', ''],
            ['5', 'Capital', '', '', '', '', ''],
            ['6', 'Reserves and Surplus', '', '', '', '', ''],
            ['7', 'Unsecured loan', '', '', '', '', ''],
            ['8', 'Net Worth (incl. GoI Subsidy as Quasi-equity)', '', '', '', '', ''],
            ['9', 'Income', '', '', '', '', ''],
            ['10', 'Gross profit', '', '', '', '', ''],
            ['11', 'Depreciation', '', '', '', '', ''],
            ['12', 'Profit after tax', '', '', '', '', ''],
            ['13', 'Gross Cash Accruals', '', '', '', '', ''],
        ]
        performance_data = [
            ['Particulars', 'Before Intervention\nQty. / Outcome', 'After Intervention\nQty. / Outcome'],
            ['Units (including details of SC/ST/Women/Minorities)', '', ''],
            ['Employment', '', ''],
            ['Production', '', ''],
            ['Exports', '', ''],
            ['Import Substitution', '', ''],
            ['Number of patent expected aimed', '', ''],
            ['Investment', '', ''],
            ['Turnover', '', ''],
            ['Profit', '', ''],
            ['Quality Certification', '', ''],
            ['Any others (No. of ZED certified units)', '', ''],
        ]
    else:
        impl_data = get_implementation_data()
        manpower_data = get_manpower_data()
        schedule_data = get_schedule_data()
        cost_data = get_cost_data()
        machinery_data = get_machinery_data()
        financing_data = get_financing_data()
        financial_data = get_financial_data()
        performance_data = get_performance_data()
    
    # PDF filename with project name and timestamp
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    safe_project_name = project_name.replace(" ", "_")
    filename = os.path.join(downloads_folder, f"DPR_{safe_project_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
    
    # PDF document setup
    doc = SimpleDocTemplate(filename, pagesize=A4,
                           rightMargin=40, leftMargin=40,
                           topMargin=30, bottomMargin=25)
    
    story = []
    styles = getSampleStyleSheet()
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=13,
        textColor=colors.black,
        spaceAfter=3,
        spaceBefore=1,
        alignment=TA_CENTER,
        bold=True,
        leading=14
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=11,
        textColor=colors.black,
        spaceAfter=2,
        spaceBefore=1,
        bold=True,
        leading=12
    )
    
    subheading_style = ParagraphStyle(
        'CustomSubHeading',
        parent=styles['Heading3'],
        fontSize=10,
        textColor=colors.black,
        spaceAfter=2,
        spaceBefore=1,
        bold=True,
        leading=11
    )
    
    normal_style = ParagraphStyle(
        'CustomNormal',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.black,
        spaceAfter=1,
        spaceBefore=0,
        leading=10
    )
    
    # Header
    story.append(Paragraph("New Guidelines MSE-CDP Page 18 of 44", normal_style))
    story.append(Spacer(1, 3))
    
    # Title
    story.append(Paragraph("Annexure-3", title_style))
    story.append(Paragraph("Format of Detailed Proposal for CFC", title_style))
    story.append(Spacer(1, 4))
    
    # Section 1
    story.append(Paragraph("1. Proposal under consideration", heading_style))
    story.append(Spacer(1, 2))
    
    # Section 2
    story.append(Paragraph("2. Brief particulars of the proposal", heading_style))
    story.append(Spacer(1, 2))
    
    section2_table = Table(section2_data, colWidths=[2.3*inch, 4.5*inch])
    section2_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(section2_table)
    story.append(Spacer(1, 3))
    
    # Section 2.1
    story.append(Paragraph("2.1. Introduction: brief about", subheading_style))
    story.append(Paragraph("2.1.1. General scenario of industrial growth/ cluster development in the state", normal_style))
    story.append(Paragraph("2.1.2. Sector for which CFC is proposed to be set up", normal_style))
    story.append(Paragraph("2.1.3. Cluster and its products, future prospects of products, Competition scenario, Backward and forward linkages", normal_style))
    story.append(Paragraph("Basic data of cluster (Number of units, type of units [Micro/Small/Medium], employment [direct /indirect], turnover, exports, etc):", normal_style))
    story.append(Paragraph("2.1.4. How the proposed CFC is relevant to the growth of the concerned cluster/ sector", normal_style))
    story.append(Spacer(1, 4))
    
    # Section 3
    story.append(Paragraph("3. Information about SPV", heading_style))
    story.append(Spacer(1, 2))
    
    spv_table = Table(spv_data, colWidths=[0.5*inch, 2.8*inch, 3.5*inch])
    spv_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(spv_table)
    story.append(Spacer(1, 4))
    
    # Continue SPV table
    spv_table2 = Table(spv_data2, colWidths=[0.5*inch, 2.8*inch, 3.5*inch])
    spv_table2.setStyle(TableStyle([
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(spv_table2)
    story.append(Spacer(1, 4))
    
    # Section 4
    story.append(Paragraph("4. Details of Project Promoters /Sponsors", heading_style))
    story.append(Paragraph("(i) Brief bio-data of Promoters", normal_style))
    story.append(Paragraph("(ii) The details of the promoters are as under:", normal_style))
    story.append(Spacer(1, 4))
    
    promoters_table = Table(promoters_data, colWidths=[1.2*inch, 0.4*inch, 0.8*inch, 0.8*inch, 1.0*inch, 1.1*inch, 0.6*inch, 0.5*inch])
    promoters_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))

    story.append(promoters_table)
    story.append(Spacer(1, 4))
    
    # Additional details
    story.append(Paragraph("(i) Brief about Compliance with KYC guidelines", normal_style))
    story.append(Paragraph("(ii) Details of connected lending - Whether the directors / promoters of SPV are having any directorship on any bank etc.", normal_style))
    story.append(Paragraph("(iii) Adverse auditors remarks, if any to be culled out from audit report, in case available. If SPV is new, it can be indicated as not applicable", normal_style))
    story.append(Paragraph("(iv) Particulars of previous assistance from financial institutions / banks - If SPV is new, it can be indicated as not applicable", normal_style))
    story.append(Paragraph("(v) Pending court cases initiated by other banks/FIs, if any - If SPV is new, it can be indicated as not applicable", normal_style))
    story.append(Paragraph("(vi) Management Set-up", normal_style))
    story.append(Paragraph("(vii) To indicate details regarding who will be the main persons involved in running of CFC, its operations etc.", normal_style))
    story.append(Spacer(1, 4))
    
    # Section 5 - Eligibility
    story.append(Paragraph("5. Eligibility as per guidelines of MSE-CDP", heading_style))
    story.append(Spacer(1, 4))
    
    eligibility_data = [
        ['S.\nNo.', 'Eligibility Criteria', 'Comments'],
        ['1.', Paragraph('The GoI grant will be restricted to 60% / 70% / 80% of the cost of Project of maximum Rs.30.00 crore as per the Scheme guidelines.', normal_style), ''],
        ['2.', Paragraph('Cost of project includes cost of Land (subject to max. of 25% of Project Cost), building, pre-operative expenses, preliminary expenses, machinery & equipment, miscellaneous fixed assets, support infrastructure such as water supply, electricity and margin money for working capital.', normal_style), ''],
        ['3.', Paragraph('The entire cost of land and building for CFC shall be met by SPV/State Government concerned.', normal_style), ''],
        ['4.', Paragraph('In case existing land and building is provided by stakeholders, the cost of land and building will be decided on the basis of valuation report prepared by an approved agency of Central/State Govt. Departments/FIs/Public Sector Banks. Cost of land and building may be taken towards contribution for the project.', normal_style), ''],
        ['5.', Paragraph('CFC can be set up in leased premises. However, the lease should be legally tenable and for a fairly long duration (say 15 years).', normal_style), ''],
        ['6.', Paragraph('Escalation in the cost of project above the sanctioned amount, due to any reason, will be borne by the SPV/ State Government. The Central Government shall not accept any financial liability arising out of operation of any CFC.', normal_style), ''],
        ['7.', Paragraph('DPR should be appraised by a bank (if bank financing is involved) / independent Technical Consultancy Organization/ SIDBI.', normal_style), ''],
        ['8.', Paragraph('Proposals approved and forwarded by the concerned state government.', normal_style), ''],
        ['9.', Paragraph('Evidence should be furnished with regard to SPV members ability to utilize at least 60% of installed capacity.', normal_style), ''],
    ]
    
    eligibility_table = Table(eligibility_data, colWidths=[0.4*inch, 5.5*inch, 0.9*inch])
    eligibility_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(eligibility_table)
    story.append(Spacer(1, 3))
    
    # Section 6
    story.append(Paragraph("6. Implementing Arrangements", heading_style))
    story.append(Spacer(1, 2))
    
    impl_table = Table(impl_data, colWidths=[4.6*inch, 2.2*inch])
    impl_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(impl_table)
    story.append(Spacer(1, 3))
    
    # Remaining sections
    story.append(Paragraph("7. Management and shareholding details:", heading_style))
    story.append(Spacer(1, 3))
    
    story.append(Paragraph("8. Technical Aspects:", heading_style))
    story.append(Paragraph("(i) Scope of the project (including components/ sections of CFC)", normal_style))
    story.append(Paragraph("(ii) Locational details and availability of infrastructural facilities", normal_style))
    story.append(Paragraph("(iii) Technology", normal_style))
    story.append(Paragraph("(iv) Provision for Industry 4.0 of AI and innovations if any", normal_style))
    story.append(Paragraph("(v) Raw materials / components", normal_style))
    story.append(Paragraph("(vi) Utilities", normal_style))
    story.append(Paragraph("    (a) Power", normal_style))
    story.append(Paragraph("    (b) Water", normal_style))
    story.append(Paragraph("(vii) Effluent disposal", normal_style))
    story.append(Paragraph("(viii) Manpower", normal_style))
    story.append(Paragraph("The details of the manpower are as under:", normal_style))
    story.append(Spacer(1, 2))
    
    manpower_table = Table(manpower_data, colWidths=[0.7*inch, 4.3*inch, 1.8*inch])
    manpower_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(manpower_table)
    story.append(Spacer(1, 4))
    
    # Section 9
    story.append(Paragraph("9. Implementation Schedule:", heading_style))
    story.append(Spacer(1, 2))
    
    schedule_table = Table(schedule_data, colWidths=[3.2*inch, 1.8*inch, 1.8*inch])
    schedule_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(schedule_table)
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>Note:</b> PERT Chart for all activities to be accomplished in accordance with activity-wise time line as prescribed in MSE-CDP guidelines will mandatory be a part of DPR.", normal_style))
    story.append(Spacer(1, 4))
    
    # Section 10
    story.append(Paragraph("10. Project components:", heading_style))
    story.append(Paragraph("(i) Estimated Project Cost (Rs. in lakh):", subheading_style))
    story.append(Spacer(1, 2))
    
    cost_table = Table(cost_data, colWidths=[0.7*inch, 4.3*inch, 1.8*inch])
    cost_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(cost_table)
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("(ii) Details of Land, Site Development and Building & Civil Work", normal_style))
    story.append(Spacer(1, 3))
    
    story.append(Paragraph("(iii) Plant & Machinery:", normal_style))
    story.append(Paragraph("(Rs. in lakh)", normal_style))
    story.append(Spacer(1, 2))
    
    machinery_table = Table(machinery_data, colWidths=[0.7*inch, 3.4*inch, 1*inch, 1.7*inch])
    machinery_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(machinery_table)
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("(iv) Comments on Plant and Machineries from O/o DC, MSME:", normal_style))
    story.append(Paragraph("(v) Misc. fixed assets", normal_style))
    story.append(Paragraph("(vi) Preliminary expenses", normal_style))
    story.append(Paragraph("(vii) Pre-operative expenses", normal_style))
    story.append(Paragraph("(viii) Contingency Provisions:", normal_style))
    story.append(Paragraph("(ix) Margin money for Working Capital", normal_style))
    story.append(Spacer(1, 3))
    
    story.append(Paragraph("(x) Proposed Means of Financing:", normal_style))
    story.append(Paragraph("(Rs. in lakh)", normal_style))
    story.append(Spacer(1, 2))
    
    financing_table = Table(financing_data, colWidths=[0.7*inch, 3.2*inch, 1.2*inch, 1.7*inch])
    financing_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(financing_table)
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("(xi) SPV contribution:", normal_style))
    story.append(Paragraph("(xii) Grant-in-aid from Govt. of India under MSE-CDP", normal_style))
    story.append(Paragraph("(xiii) Grant-in-aid from the State Government", normal_style))
    story.append(Paragraph("(xiv) Bank Loan/ others", normal_style))
    story.append(Paragraph("(xv) Arrangements for utilization of facilities by cluster units:", normal_style))
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("11. Fund requirement / availability analysis: The details must be provided keeping in view that pace of the project is not suffered due to non-availability of funds in time.", heading_style))
    story.append(Spacer(1, 3))
    
    story.append(Paragraph("12. Usage Charges:", heading_style))
    story.append(Spacer(1, 3))

    story.append(Paragraph("13. Comments on Commercial viability:", heading_style))
    story.append(Spacer(1, 4))
    
    story.append(Paragraph("14. Financial Economic viability:", heading_style))
    story.append(Paragraph("Assumptions underlying the profitability estimates, projected cash flow statements and projected balance sheet are placed at Annexure and the summary of key parameters for the first 5 years are given below:-", normal_style))
    story.append(Paragraph("                           (Rs. in lakh)", normal_style))
    story.append(Spacer(1, 2))
    
    financial_table = Table(financial_data, colWidths=[0.4*inch, 2.4*inch, 0.7*inch, 0.7*inch, 0.7*inch, 0.7*inch, 0.7*inch])
    financial_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 7.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(financial_table)
    story.append(Spacer(1, 3))
    
    story.append(Paragraph("The projected revenue of SPV is based upon the following major assumptions:", normal_style))
    story.append(Spacer(1, 6))
    
    # Section 15
    story.append(Paragraph("15. Projected performance of the cluster after proposed intervention (in terms of production, domestic sales / exports and direct, indirect employment, etc.)", heading_style))
    story.append(Spacer(1, 2))
    
    performance_table = Table(performance_data, colWidths=[2.3*inch, 2.3*inch, 2.3*inch])
    performance_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8.5),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 2),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    
    story.append(performance_table)
    story.append(Spacer(1, 4))
    
    # Section 16
    story.append(Paragraph("16. Status of Government approvals", heading_style))
    story.append(Paragraph("(i) Pollution control", normal_style))
    story.append(Paragraph("(ii) Permission for land use (conversion for industrial purpose)", normal_style))
    story.append(Spacer(1, 4))
    
    # Section 17
    story.append(Paragraph("17. Favorable and Risk Factors of the project : SWOT Analysis", heading_style))
    story.append(Spacer(1, 4))
    
    # Section 18
    story.append(Paragraph("18. Risk Mitigation Framework", heading_style))
    story.append(Paragraph("Key risks during the implementation and operations phase of the Project and the mitigations measures thereof could be as below:", normal_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>During implementation:</b>", normal_style))
    story.append(Spacer(1, 3))
    story.append(Paragraph("<b>During operations:</b>", normal_style))
    story.append(Spacer(1, 4))
    
    # Section 19
    story.append(Paragraph("19. Economics of the project", heading_style))
    story.append(Paragraph("(a) Debt Service coverage ratio (Projections for 10 years)", normal_style))
    story.append(Spacer(1, 2))
    
    # DSCR Formula
    dscr_formula = """<para align="center">
    DSCR = (Net Profit + Interest(TL) + Depreciation) / (installment(TL) + Interest(TL))
    </para>"""
    story.append(Paragraph(dscr_formula, normal_style))
    story.append(Spacer(1, 3))
    
    story.append(Paragraph("(b) Balance sheet & P/L account (projection for 10 years)", normal_style))
    story.append(Spacer(1, 3))
    
    # Break Even Formula
    breakeven_formula = """<para align="left">
    (c) Break Even Point = Fixed Cost / Contribution on Sales (Sales - Variable Cost)
    </para>"""
    story.append(Paragraph(breakeven_formula, normal_style))
    story.append(Spacer(1, 4))
    
    # Section 20
    story.append(Paragraph("20. Commercial Viability: Following financial appraisal tools will be employed for assessing commercial viability of the Project:", heading_style))
    story.append(Spacer(1, 2))
    
    story.append(Paragraph("<b>(i) Return on Capital Employed (ROCE):</b> The total return generated by the project over its entire projected life will be averaged to find out the average yearly return. The simple acceptance rule for the investment is that the return (incorporating benefit of grant-in-aid assistance) is sufficiently larger than the interest on capital employed. Return in excess of 25% is desirable.", normal_style))
    story.append(Spacer(1, 2))
    
    story.append(Paragraph("<b>(ii) Debt Service Coverage Ratio:</b> Acceptance rule will be cumulative DSCR of 3:1 during repayment period.", normal_style))
    story.append(Spacer(1, 2))
    
    story.append(Paragraph("<b>(iii) Break-Even (BE) Analysis:</b> Break-even point should be below 60 per cent of the installed capacity.", normal_style))
    story.append(Spacer(1, 2))
    
    story.append(Paragraph("<b>(iv) Sensitivity Analysis:</b> Sensitivity analysis will be pursued for all the major financial parameters/indicators in terms of a 5-10 per cent drop in user charges or fall in capacity utilisation by 10-20 per cent.", normal_style))
    story.append(Spacer(1, 2))
    
    story.append(Paragraph("<b>(v) Net Present Value (NPV):</b> Net Present Value of the Project needs to be positive and the Internal Rate of return (IRR) should be above 10 per cent. The rate of discount to be adopted for estimation of NPV will be 10 per cent. The Project life may be considered to be a maximum of 10 years. The life of the Project to be considered for this purpose needs to be supported by recommendation of a technical expert/institution.", normal_style))
    story.append(Spacer(1, 4))
    
    # Section 21
    story.append(Paragraph("21. Conclusion", heading_style))
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("*****", ParagraphStyle('Center', parent=normal_style, alignment=TA_CENTER)))
    
    # Build PDF
    doc.build(story)
    
    print(f"\n✓✓✓ PDF successfully created!")
    print(f"✓ Project Name: {project_name}")
    print(f"✓ Filename: {os.path.basename(filename)}")
    print(f"✓ Location: {downloads_folder}")
    print(f"✓ Full path: {filename}")
    return filename

# Run the function
if __name__ == "__main__":
    try:
        print("\n" + "="*60)
        print("  DPR TEMPLATE PDF GENERATOR - INTERACTIVE MODE")
        print("="*60)
        
        # Get basic inputs first
        print("\n=== Basic Project Information ===\n")
        
        project_name = input("Project Name: ").strip()
        candidate_name = input("Candidate/SPV Name: ").strip()
        address = input("Address: ").strip()
        
        # Validate inputs
        if not project_name or not candidate_name or not address:
            print("\n✗ Error: Project Name, Candidate Name, and Address are required!")
            exit(1)
        
        print("\n✓ Basic information captured!")
        print("\n" + "="*60)
        print("Now you'll be asked to fill table data.")
        print("Press Enter to skip any field.")
        print("Type 'done' at ANY point to stop and generate PDF immediately.")
        print("="*60)
        
        # Generate PDF with interactive data collection
        pdf_file = create_dpr_pdf(project_name, candidate_name, address)
        
        print("\n" + "="*60)
        print("  ✓✓✓ PDF CREATION SUCCESSFUL!")
        print("="*60 + "\n")
    except KeyboardInterrupt:
        print("\n\n✗ Process interrupted by user.")
    except Exception as e:
        print(f"\n✗ Error occurred: {e}")
        import traceback
        traceback.print_exc()