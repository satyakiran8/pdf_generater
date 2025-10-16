import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib import colors
from datetime import datetime

def create_dpr_pdf(project_name, candidate_name, address, **kwargs):
    """
    Generate DPR PDF with dynamic data filling in both tables and text
    
    Usage:
    create_dpr_pdf(
        "My Project",
        "John Doe", 
        "123 Main St",
        Location_of_Common_Facility_Centre="Industrial Area, Phase 2",
        Profit="50 Lakhs",
        Age_years=["35", "42", "38"],
    )
    """
    
    # Track which keys have been used
    used_keys = set()
    
    # Helper function to find and fill table values
    def fill_table_data(table_data, kwargs):
        """Fill table with values from kwargs based on key matching"""
        
        for row_idx, row in enumerate(table_data):
            for col_idx, cell in enumerate(row):
                if isinstance(cell, str) and cell.strip():
                    # Clean the cell text for matching
                    cell_key = cell.strip().replace('\n', ' ').replace('  ', ' ')
                    
                    best_match = None
                    best_match_score = 0
                    best_value = None
                    
                    # First check for exact text match (for standalone text in tables)
                    exact_key, exact_value = find_text_match(cell_key, kwargs)
                    if exact_key and exact_value:
                        # Found exact match - fill next column
                        if isinstance(exact_value, list):
                            for i, val in enumerate(exact_value):
                                if col_idx + 1 + i < len(row):
                                    table_data[row_idx][col_idx + 1 + i] = str(val)
                        else:
                            if col_idx + 1 < len(row):
                                table_data[row_idx][col_idx + 1] = str(exact_value)
                        used_keys.add(exact_key)
                        continue
                    
                    # If no exact match, try fuzzy matching (existing logic)
                    # Check all kwargs for matching keys
                    for key, value in kwargs.items():
                        if key in used_keys:
                            continue  # Skip already used keys
                        
                        # Create searchable version of key
                        search_key = key.replace('_', ' ').lower().strip()
                        cell_lower = cell_key.lower().strip()
                        
                        # Remove common prefixes for better matching
                        cell_cleaned = cell_lower
                        for prefix in ['a. ', 'b. ', 'c. ', 'd. ', 'e. ', 'f. ', 'g. ', 'h. ', 'i. ', 'j. ', '(i)', '(ii)', '(iii)', '(iv)', '(v)', '(vi)', '(vii)', '(viii)', '(ix)', '(x)', '(xi)', '(xii)', '(xiii)', '(xiv)', '(xv)']:
                            if cell_cleaned.startswith(prefix):
                                cell_cleaned = cell_cleaned[len(prefix):].strip()
                                break
                        
                        # Calculate match score (higher = better match)
                        match_score = 0
                        
                        # Exact match with cleaned cell gets highest score
                        if search_key == cell_cleaned:
                            match_score = 10000
                        # Exact match with original cell
                        elif search_key == cell_lower:
                            match_score = 9000
                        # Cell cleaned contains the full search key
                        elif search_key in cell_cleaned and len(search_key) > 5:
                            # Only if search key is significant portion of cell
                            ratio = len(search_key) / len(cell_cleaned)
                            if ratio > 0.5:  # Search key is at least 50% of cell text
                                match_score = 5000 + int(ratio * 1000)
                        # All words from search key present in cell
                        else:
                            search_words = set(search_key.split())
                            cell_words = set(cell_cleaned.split())
                            
                            # Must have at least 2 words to consider word matching
                            if len(search_words) >= 2:
                                common_words = search_words.intersection(cell_words)
                                # All search words must be present
                                if len(common_words) == len(search_words):
                                    match_score = 3000 + len(common_words) * 100
                                # Most search words present (at least 70%)
                                elif len(common_words) >= len(search_words) * 0.7:
                                    match_score = 1000 + len(common_words) * 50
                        
                        # Update best match if this is better
                        if match_score > best_match_score:
                            best_match_score = match_score
                            best_match = key
                            best_value = value
                    
                    # If we found a good match (score > 1000), fill it
                    if best_match and best_match_score > 1000:
                        # Handle list values (multiple columns)
                        if isinstance(best_value, list):
                            # Fill multiple columns with list values
                            for i, val in enumerate(best_value):
                                if col_idx + 1 + i < len(row):
                                    table_data[row_idx][col_idx + 1 + i] = str(val)
                        else:
                            # Fill next column with single value
                            if col_idx + 1 < len(row):
                                table_data[row_idx][col_idx + 1] = str(best_value)
                        
                        # Mark this key as used
                        used_keys.add(best_match)
                        break  # Move to next row after filling
        
        return table_data
    
    # Helper function to check if text matches a key exactly
    def find_text_match(text, kwargs):
        """Find exact match for text in kwargs keys"""
        if not isinstance(text, str):
            return None, None
            
        text_clean = text.strip().lower()
        
        # Remove common prefixes/suffixes and section numbers
        prefixes_to_remove = [
            '(i) ', '(ii) ', '(iii) ', '(iv) ', '(v) ', '(vi) ', 
            '(vii) ', '(viii) ', '(ix) ', '(x) ', '(xi) ', '(xii) ', 
            '(xiii) ', '(xiv) ', '(xv) ', '(xvi) ', '(xvii) ', '(xviii) ',
            'a. ', 'b. ', 'c. ', 'd. ', 'e. ', 'f. ', 'g. ', 'h. ', 'i. ', 'j. ',
            '1. ', '2. ', '3. ', '4. ', '5. ', '6. ', '7. ', '8. ', '9. ',
            '10. ', '11. ', '12. ', '13. ', '14. ', '15. ', '16. ', '17. ', '18. ', '19. ', '20. ', '21. '
        ]
        
        for prefix in prefixes_to_remove:
            if text_clean.startswith(prefix.lower()):
                text_clean = text_clean[len(prefix):].strip()
                break
        
        # Remove trailing colons
        text_clean = text_clean.rstrip(':').strip()
        
        # Check for exact match with any key
        for key, value in kwargs.items():
            if key in used_keys:
                continue
                
            key_clean = key.replace('_', ' ').lower().strip()
            # Also remove trailing colons from key
            key_clean = key_clean.rstrip(':').strip()
            
            # Exact match required
            if text_clean == key_clean:
                return key, value
        
        return None, None
    
    # PDF filename with project name and timestamp
    downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")
    
    # Remove invalid Windows filename characters
    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
    safe_project_name = project_name
    for char in invalid_chars:
        safe_project_name = safe_project_name.replace(char, '')
    
    # Replace spaces with underscores and limit length
    safe_project_name = safe_project_name.replace(" ", "_")[:50]
    
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
    
    value_style = ParagraphStyle(
        'ValueStyle',
        parent=styles['Normal'],
        fontSize=9,
        textColor=colors.blue,
        spaceAfter=1,
        spaceBefore=0,
        leading=10,
        leftIndent=15
    )
    
    # Helper to add text with optional value
    def add_text(text, style=normal_style):
        """Add text and check if it needs a value filled below it"""
        story.append(Paragraph(text, style))
        
        # Check if this text matches any key
        matched_key, value = find_text_match(text, kwargs)
        if matched_key and value:
            # Add value below the matched text
            if isinstance(value, list):
                value_text = ", ".join(str(v) for v in value)
            else:
                value_text = str(value)
            
            story.append(Paragraph(f"→ {value_text}", value_style))
            used_keys.add(matched_key)
    
    # Header
    story.append(Paragraph("New Guidelines MSE-CDP Page 18 of 44", normal_style))
    story.append(Spacer(1, 3))
    
    # Title
    story.append(Paragraph("Annexure-3", title_style))
    story.append(Paragraph("Format of Detailed Proposal for CFC", title_style))
    story.append(Spacer(1, 4))
    
    # Section 1
    add_text("1. Proposal under consideration", heading_style)
    story.append(Spacer(1, 2))
    
    # Section 2
    add_text("2. Brief particulars of the proposal", heading_style)
    story.append(Spacer(1, 2))
    
    # Section 2 Table
    section2_data = [
        ['Name of applicant,\ncontact details, etc', 'As CFC Registered address / administrative address may be different\nfrom CFC facilities address, the same may be provided'],
        ['Location of Common\nFacility Centre', 'Address where facilities are proposed may be provided'],
        ['Main facilities being\nproposed', 'Details of facilities to be provided']
    ]
    
    # Fill with kwargs data
    section2_data = fill_table_data(section2_data, kwargs)
    
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
    add_text("2.1. Introduction: brief about", subheading_style)
    add_text("2.1.1. General scenario of industrial growth/ cluster development in the state")
    add_text("2.1.2. Sector for which CFC is proposed to be set up")
    add_text("2.1.3. Cluster and its products, future prospects of products, Competition scenario, Backward and forward linkages")
    add_text("Basic data of cluster (Number of units, type of units [Micro/Small/Medium], employment [direct /indirect], turnover, exports, etc):")
    add_text("2.1.4. How the proposed CFC is relevant to the growth of the concerned cluster/ sector")
    story.append(Spacer(1, 4))
    
    # Section 3
    add_text("3. Information about SPV", heading_style)
    story.append(Spacer(1, 2))
    
    # Combine candidate name and address
    name_address = f"{candidate_name}, {address}"
    
    # SPV Information Table
    spv_data = [
        ['S. No.', 'Description', 'Details/ Compliance'],
        ['(i)', 'Name and address', name_address],
        ['(ii)', 'Registration details of SPV\n(including registration as Section\n8 company under the Companies\nAct 2013)', ''],
        ['(iii)', 'Names of the State Govt and\nMSME officials in SPV', ''],
        ['(iv)', 'Date of formation of the\ncompany', ''],
        ['(v)', 'Date of commencement of\nbusiness', ''],
        ['(vi)', 'Number of MSE Member Units', ''],
        ['(vii)', 'Bye laws or MoA and AoA\nsubmitted', ''],
        ['(viii)', 'Main objectives of the SPV', ''],
        ['(ix)', 'SPV to have a character of\ninclusiveness wherein provision\nfor enrolling new members to\nenable prospective entrepreneurs\nin the cluster to utilise the\nfacility', ''],
        ['(x)', "Clause about 'Profits/surplus\n to be ploughed back\n to CFC' included or not", ''],
        ['(xi)', 'Authorized share capital', ''],
        ['(xii)', 'Shareholding Pattern\n(Annexure-3 to be filled in)', ''],
    ]
    
    spv_data = fill_table_data(spv_data, kwargs)
    
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
    spv_data2 = [
        ['(xiii)', 'Commitment letter for SPV\nUpfront contribution', ''],
        ['(xiv)', 'Project specific A/c in schedule\nA bank', ''],
        ['(xv)', "Clause about 'CFC may be\n utilised by SPV members as also\nothers in a cluster and \nEvidence for SPV members'\nability to utilise at least 60% of installed capacity'", ''],
        ['(xvi)', 'Main Role of SPV', ''],
        ['(xvii)', 'Trust building of SPV so that\nCFC may be successful', ''],
    ]
    
    spv_data2 = fill_table_data(spv_data2, kwargs)
    
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
    add_text("4. Details of Project Promoters /Sponsors", heading_style)
    add_text("(i) Brief bio-data of Promoters")
    add_text("(ii) The details of the promoters are as under:")
    story.append(Spacer(1, 4))
    
    # Promoters Table
    promoters_data = [
        ['Name of the\nOffice bearers\n of the SPV', '', '', '', '', '', '', ''],
        ['Age (years)', '', '', '', '', '', '', ''],
        ['Educational\n Qualification', '', '', '', '', '', '', ''],
        ['Relationship with \nthe chief promoter', '', '', '', '', '', '', ''],
        ['Experience in what\n capacity/ industry/ years', '', '', '', '', '', '', ''],
        ['Income Tax / Wealth \nTax Status \n(returns for 3 years\n to be furnished)', '', '', '', '', '', '', ''],
        ['Other concerns \ninterest / in which capacity \n/financial stake', '', '', '', '', '', '', '']
    ]
    
    promoters_data = fill_table_data(promoters_data, kwargs)

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
    add_text("(i) Brief about Compliance with KYC guidelines")
    add_text("(ii) Details of connected lending - Whether the directors / promoters of SPV are having any directorship on any bank etc.")
    add_text("(iii) Adverse auditors remarks, if any to be culled out from audit report, in case available. If SPV is new, it can be indicated as not applicable")
    add_text("(iv) Particulars of previous assistance from financial institutions / banks - If SPV is new, it can be indicated as not applicable")
    add_text("(v) Pending court cases initiated by other banks/FIs, if any - If SPV is new, it can be indicated as not applicable")
    add_text("(vi) Management Set-up")
    add_text("(vii) To indicate details regarding who will be the main persons involved in running of CFC, its operations etc.")
    story.append(Spacer(1, 4))
    
    # Section 5 - Eligibility
    add_text("5. Eligibility as per guidelines of MSE-CDP", heading_style)
    story.append(Spacer(1, 4))
    
    # Eligibility data
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
    add_text("6. Implementing Arrangements", heading_style)
    story.append(Spacer(1, 2))
    
    impl_data = [
        ['Description', 'Compliance'],
        ['a. Name of Implementation Agency', ''],
        ['b. Role of Implementing Agency (e.g. implementation and monitoring of project,\nsending regular progress reports, issuing proper UCs, )', ''],
        ['c. Implementation Period', ''],
        ['d. Commitment of State Government upfront contribution', ''],
        ['e. Commitment of Loans (Working capital and/ or term loan)', ''],
    ]
    
    impl_data = fill_table_data(impl_data, kwargs)
    
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
    add_text("7. Management and shareholding details:", heading_style)
    story.append(Spacer(1, 3))
    
    add_text("8. Technical Aspects:", heading_style)
    add_text("(i) Scope of the project (including components/ sections of CFC)")
    add_text("(ii) Locational details and availability of infrastructural facilities")
    add_text("(iii) Technology")
    add_text("(iv) Provision for Industry 4.0 of AI and innovations if any")
    add_text("(v) Raw materials / components")
    add_text("(vi) Utilities")
    add_text("    (a) Power")
    add_text("    (b) Water")
    add_text("(vii) Effluent disposal")
    add_text("(viii) Manpower")
    add_text("The details of the manpower are as under:")
    story.append(Spacer(1, 2))
    
    manpower_data = [
        ['S. No.', 'Description of the employee', 'Number'],
        ['1', '', ''],
        ['2', '', ''],
        ['3', '', ''],
        ['4', '', ''],
    ]
    
    manpower_data = fill_table_data(manpower_data, kwargs)
    
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
    add_text("9. Implementation Schedule:", heading_style)
    story.append(Spacer(1, 2))
    
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
    
    schedule_data = fill_table_data(schedule_data, kwargs)
    
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
    add_text("10. Project components:", heading_style)
    add_text("(i) Estimated Project Cost (Rs. in lakh):", subheading_style)
    story.append(Spacer(1, 2))
    
    cost_data = [
        ['S. No.', 'Particulars', 'Amount'],
        ['1', 'Land and Building', ''],
        ['2', 'Plant & Machinery including MFA,\nInstallation, Taxes/duties,\nContingencies, etc.', ''],
        ['3', 'Preliminary & Pre-operative expenses', ''],
        ['4', 'Margin money for Working Capital', ''],
        ['', 'Total', ''],
    ]
    
    cost_data = fill_table_data(cost_data, kwargs)
    
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
    
    add_text("(ii) Details of Land, Site Development and Building & Civil Work")
    story.append(Spacer(1, 3))
    
    add_text("(iii) Plant & Machinery:")
    story.append(Paragraph("(Rs. in lakh)", normal_style))
    story.append(Spacer(1, 2))
    
    machinery_data = [
        ['S. No.', 'Description', 'No.', 'Amount'],
        ['1', '', '', ''],
        ['2', '', '', ''],
        ['3', '', '', ''],
        ['4', '', '', ''],
    ]
    
    machinery_data = fill_table_data(machinery_data, kwargs)
    
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
    
    add_text("(iv) Comments on Plant and Machineries from O/o DC, MSME:")
    add_text("(v) Misc. fixed assets")
    add_text("(vi) Preliminary expenses")
    add_text("(vii) Pre-operative expenses")
    add_text("(viii) Contingency Provisions:")
    add_text("(ix) Margin money for Working Capital")
    story.append(Spacer(1, 3))
    
    add_text("(x) Proposed Means of Financing:")
    story.append(Paragraph("(Rs. in lakh)", normal_style))
    story.append(Spacer(1, 2))
    
    financing_data = [
        ['S. No.', 'Particulars', 'Percentage', 'Amount'],
        ['', '', '', ''],
        ['', 'Total', '', ''],
    ]
    
    financing_data = fill_table_data(financing_data, kwargs)
    
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
    
    add_text("(xi) SPV contribution:")
    add_text("(xii) Grant-in-aid from Govt. of India under MSE-CDP")
    add_text("(xiii) Grant-in-aid from the State Government")
    add_text("(xiv) Bank Loan/ others")
    add_text("(xv) Arrangements for utilization of facilities by cluster units:")
    story.append(Spacer(1, 4))
    
    add_text("11. Fund requirement / availability analysis: The details must be provided keeping in view that pace of the project is not suffered due to non-availability of funds in time.", heading_style)
    story.append(Spacer(1, 3))
    
    add_text("12. Usage Charges:", heading_style)
    story.append(Spacer(1, 3))

    add_text("13. Comments on Commercial viability:", heading_style)
    story.append(Spacer(1, 4))
    
    add_text("14. Financial Economic viability:", heading_style)
    story.append(Paragraph("Assumptions underlying the profitability estimates, projected cash flow statements and projected balance sheet are placed at Annexure and the summary of key parameters for the first 5 years are given below:-", normal_style))
    story.append(Paragraph("                           (Rs. in lakh)", normal_style))
    story.append(Spacer(1, 2))
    
    financial_data = [
        ['S. No.', 'Particulars', 'FY 1', 'FY 2', 'FY3', 'FY4', 'FY5'],
        ['1', 'Net Block', '', '', '', '', ''],
        ['2', 'Current Assets (incl.\ncash/bank balance)', '', '', '', '', ''],
        ['3', 'Current Liabilities (incl.\nprincipal installment falling\ndue during the year)', '', '', '', '', ''],
        ['4', 'Long term borrowings', '', '', '', '', ''],
        ['5', 'Capital', '', '', '', '', ''],
        ['6', 'Reserves and Surplus', '', '', '', '', ''],
        ['7', 'Unsecured loan', '', '', '', '', ''],
        ['8', 'Net Worth (incl. GoI Subsidy\nas Quasi-equity)', '', '', '', '', ''],
        ['9', 'Income', '', '', '', '', ''],
        ['10', 'Gross profit', '', '', '', '', ''],
        ['11', 'Depreciation', '', '', '', '', ''],
        ['12', 'Profit after tax', '', '', '', '', ''],
        ['13', 'Gross Cash Accruals', '', '', '', '', ''],
    ]
    
    financial_data = fill_table_data(financial_data, kwargs)
    
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
    
    add_text("The projected revenue of SPV is based upon the following major assumptions:")
    story.append(Spacer(1, 4))
    
    # Section 15
    add_text("15. Projected performance of the cluster after proposed intervention (in terms of production, domestic sales / exports and direct, indirect employment, etc.)", heading_style)
    story.append(Spacer(1, 2))
    
    performance_data = [
        ['Particulars', 'Before Intervention\nQty. / Outcome', 'After Intervention\nQty. / Outcome'],
        ['Units (including\ndetails of\nSC/ST/Women\n/Minorities)', '', ''],
        ['Employment', '', ''],
        ['Production', '', ''],
        ['Exports', '', ''],
        ['Import Substitution', '', ''],
        ['Number of patent\nexpected aimed', '', ''],
        ['Investment', '', ''],
        ['Turnover', '', ''],
        ['Profit', '', ''],
        ['Quality\nCertification', '', ''],
        ['Any others (No. of\nZED certified\nunits)', '', ''],
    ]
    
    performance_data = fill_table_data(performance_data, kwargs)
    
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
    add_text("16. Status of Government approvals", heading_style)
    add_text("(i) Pollution control")
    add_text("(ii) Permission for land use (conversion for industrial purpose)")
    story.append(Spacer(1, 4))
    
    # Section 17
    add_text("17. Favorable and Risk Factors of the project : SWOT Analysis", heading_style)
    story.append(Spacer(1, 4))
    
    # Section 18
    add_text("18. Risk Mitigation Framework", heading_style)
    story.append(Paragraph("Key risks during the implementation and operations phase of the Project and the mitigations measures thereof could be as below:", normal_style))
    story.append(Spacer(1, 3))
    add_text("During implementation:")
    story.append(Spacer(1, 3))
    add_text("During operations:")
    story.append(Spacer(1, 4))
    
    # Section 19
    add_text("19. Economics of the project", heading_style)
    add_text("(a) Debt Service coverage ratio (Projections for 10 years)")
    story.append(Spacer(1, 2))
    
    # DSCR Formula
    dscr_formula = """<para align="center">
    DSCR = (Net Profit + Interest(TL) + Depreciation) / (installment(TL) + Interest(TL))
    </para>"""
    story.append(Paragraph(dscr_formula, normal_style))
    story.append(Spacer(1, 3))
    
    add_text("(b) Balance sheet & P/L account (projection for 10 years)")
    story.append(Spacer(1, 3))
    
    # Break Even Formula
    breakeven_formula = """<para align="left">
    (c) Break Even Point = Fixed Cost / Contribution on Sales (Sales - Variable Cost)
    </para>"""
    story.append(Paragraph(breakeven_formula, normal_style))
    story.append(Spacer(1, 4))
    
    # Section 20
    add_text("20. Commercial Viability: Following financial appraisal tools will be employed for assessing commercial viability of the Project:", heading_style)
    story.append(Spacer(1, 2))
    
    add_text("(i) Return on Capital Employed (ROCE):")
    story.append(Paragraph("The total return generated by the project over its entire projected life will be averaged to find out the average yearly return. The simple acceptance rule for the investment is that the return (incorporating benefit of grant-in-aid assistance) is sufficiently larger than the interest on capital employed. Return in excess of 25% is desirable.", normal_style))
    story.append(Spacer(1, 2))
    
    add_text("(ii) Debt Service Coverage Ratio:")
    story.append(Paragraph("Acceptance rule will be cumulative DSCR of 3:1 during repayment period.", normal_style))
    story.append(Spacer(1, 2))
    
    add_text("(iii) Break-Even (BE) Analysis:")
    story.append(Paragraph("Break-even point should be below 60 per cent of the installed capacity.", normal_style))
    story.append(Spacer(1, 2))
    
    add_text("(iv) Sensitivity Analysis:")
    story.append(Paragraph("Sensitivity analysis will be pursued for all the major financial parameters/indicators in terms of a 5-10 per cent drop in user charges or fall in capacity utilisation by 10-20 per cent.", normal_style))
    story.append(Spacer(1, 2))
    
    add_text("(v) Net Present Value (NPV):")
    story.append(Paragraph("Net Present Value of the Project needs to be positive and the Internal Rate of return (IRR) should be above 10 per cent. The rate of discount to be adopted for estimation of NPV will be 10 per cent. The Project life may be considered to be a maximum of 10 years. The life of the Project to be considered for this purpose needs to be supported by recommendation of a technical expert/institution.", normal_style))
    story.append(Spacer(1, 4))
    
    # Section 21
    add_text("21. Conclusion", heading_style)
    story.append(Spacer(1, 15))
    
    story.append(Paragraph("*****", ParagraphStyle('Center', parent=normal_style, alignment=TA_CENTER)))
    
    # Build PDF
    doc.build(story)
    
    print(f"\n✓✓✓ PDF successfully created!")
    print(f"✓ Project Name: {project_name}")
    print(f"✓ Filename: {os.path.basename(filename)}")
    print(f"✓ Location: {downloads_folder}")
    print(f"✓ Full path: {filename}")
    print(f"✓ Fields filled: {len(used_keys)} out of {len(kwargs)} provided")
    
    # Show which keys were used
    if used_keys:
        print(f"\n✓ Successfully filled keys:")
        for key in sorted(used_keys):
            print(f"  - {key}")
    
    # Show unused keys
    unused = set(kwargs.keys()) - used_keys
    if unused:
        print(f"\n⚠ Unused keys (no exact match found):")
        for key in sorted(unused):
            print(f"  - {key}")
    
    return filename

# Run the function
if __name__ == "__main__":
    try:
        print("\n" + "="*60)
        print("  DPR DYNAMIC PDF GENERATOR")
        print("="*60)
        
        # Get basic inputs
        print("\n=== Basic Information ===\n")
        
        project_name = input("Project Name: ").strip()
        candidate_name = input("Candidate/SPV Name: ").strip()
        address = input("Address: ").strip()
        
        # Validate basic inputs
        if not project_name or not candidate_name or not address:
            print("\n✗ Error: Basic fields are required!")
            exit(1)
        
        print("\n✓ Basic input captured!")
        print("\n" + "="*60)
        print("  DYNAMIC DATA INPUT")
        print("="*60)
        print("\nEnter key-value pairs to fill table fields AND text sections.")
        print("Format: key=value")
        print("For multiple values: key=[value1,value2,value3]")
        print("Type 'done' when finished.\n")
        print("Examples:")
        print("  Land_and_Building=50.00")
        print("  Age_years=[35,42,38]")
        print("  Profit=50 Lakhs")
        print("  Technology=AI-based automation\n")
        
        kwargs = {}
        
        while True:
            user_input = input("Enter data (or 'done'): ").strip()
            
            if user_input.lower() == 'done':
                break
            
            if '=' not in user_input:
                print("✗ Invalid format. Use: key=value")
                continue
            
            key, value = user_input.split('=', 1)
            key = key.strip()
            value = value.strip()
            
            # Check if value is a list
            if value.startswith('[') and value.endswith(']'):
                # Parse list
                value = value[1:-1]  # Remove brackets
                value_list = [v.strip() for v in value.split(',')]
                kwargs[key] = value_list
                print(f"✓ Added: {key} = {value_list}")
            else:
                kwargs[key] = value
                print(f"✓ Added: {key} = {value}")
        
        print(f"\n✓ Total {len(kwargs)} data fields captured!")
        print("\nGenerating PDF...\n")
        
        # Generate PDF
        pdf_file = create_dpr_pdf(project_name, candidate_name, address, **kwargs)
        
        print("\n" + "="*60)
        print("  ✓✓✓ PDF CREATION SUCCESSFUL!")
        print("="*60)
        print(f"\nOpen the PDF to view your data!\n")
        
    except Exception as e:
        print(f"\n✗ Error occurred: {e}")
        import traceback
        traceback.print_exc()