import tkinter as tk
from tkinter import ttk, filedialog
from docx import Document
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.shared import OxmlElement
from docx.oxml.ns import qn


def create_education_fields(num_education):
    global edu_degree_entries, edu_major_entries, edu_institution_entries, edu_years_entries, edu_location_entries
    edu_degree_entries = []
    edu_major_entries = []
    edu_institution_entries = []
    edu_years_entries = []
    edu_location_entries = []

    for i in range(num_education):
        row_offset = 17 + i * 5 # Adjusting row based on number of fields added

        tk.Label(app, text=f"Education {i + 1} School Name:").grid(row=row_offset, column=0, sticky="e", padx=(5, 0))
        institution_entry = tk.Entry(app)
        institution_entry.grid(row=row_offset, column=1)
        edu_institution_entries.append(institution_entry)
    
        tk.Label(app, text=f"Education {i + 1} Degree Name:").grid(row=row_offset + 1, column=0, sticky="e", padx=(5, 0))
        degree_entry = tk.Entry(app)
        degree_entry.grid(row=row_offset+ 1, column=1)
        edu_degree_entries.append(degree_entry)

        tk.Label(app, text=f"Education {i + 1} Major:").grid(row=row_offset + 2, column=0, sticky="e", padx=(5, 0))
        major_entry = tk.Entry(app)
        major_entry.grid(row=row_offset + 2, column=1)
        edu_major_entries.append(major_entry)

        tk.Label(app, text=f"Education {i + 1} Start M/Y - End M/Y:").grid(row=row_offset + 3, column=0, sticky="e", padx=(5, 0))
        school_dates_entry = tk.Entry(app)
        school_dates_entry.grid(row=row_offset + 3, column=1)
        edu_years_entries.append(school_dates_entry)

        tk.Label(app, text=f"Education {i + 1} City,Country: ").grid(row=row_offset + 4, column=0, sticky="e", padx=(5, 0))
        school_location_entry = tk.Entry(app)
        school_location_entry.grid(row=row_offset + 4, column=1)
        edu_location_entries.append(school_location_entry)


def create_work_experience_fields(num_experiences):
    global work_role_entries, work_company_entries, work_dates_entries, work_location_entries
    work_role_entries = []
    work_company_entries = []
    work_dates_entries = []
    work_location_entries = []

    for i in range(num_experiences):
        row_offset = 19 + num_education * 4 + i * 4  # Adjusting row based on education fields

        tk.Label(app, text=f"Work Experience {i + 1} Company Name:").grid(row=row_offset, column=0, sticky="e", padx=(5, 0))
        company_entry = tk.Entry(app)
        company_entry.grid(row=row_offset, column=1)
        work_company_entries.append(company_entry)
    
        tk.Label(app, text=f"Work Experience {i + 1} Job Position/Title:").grid(row=row_offset + 1, column=0, sticky="e", padx=(5, 0))
        role_entry = tk.Entry(app)
        role_entry.grid(row=row_offset + 1, column=1)
        work_role_entries.append(role_entry)

        tk.Label(app, text=f"Work Experience {i + 1} Start M/Y - End M/Y:").grid(row=row_offset + 2, column=0, sticky="e", padx=(5, 0))
        work_dates_entry = tk.Entry(app)
        work_dates_entry.grid(row=row_offset + 2, column=1)
        work_dates_entries.append(work_dates_entry)

        tk.Label(app, text=f"Work Experience {i + 1} City,Country:").grid(row=row_offset + 3, column=0, sticky="e", padx=(5, 0))
        work_location_entry = tk.Entry(app)
        work_location_entry.grid(row=row_offset + 3, column=1)
        work_location_entries.append(work_location_entry)


def adjust_alignment():
    for widget in app.grid_slaves():
        widget.grid_configure(sticky="w")

def generate_fields():
    global num_education, num_experiences
    num_education = int(num_education_entry.get())
    num_experiences = int(num_experiences_entry.get())
    create_education_fields(num_education)
    create_work_experience_fields(num_experiences)
    submit_button.grid(row=39 + num_education * 4 + num_experiences * 3, column=0, columnspan=2)
    adjust_alignment()

def add_hyperlink(paragraph, text, url):
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a new run object (a wrapper over a 'w:r' element)
    new_run = docx.text.run.Run(
        docx.oxml.shared.OxmlElement('w:r'), paragraph)
    new_run.text = text

    # Set the run's style to the builtin hyperlink style, defining it if necessary
    new_run.font.color.rgb = docx.shared.RGBColor(0, 0, 255)
    new_run.font.underline = True

    # Join all the xml elements together
    hyperlink.append(new_run._element)
    paragraph._p.append(hyperlink)
    return hyperlink

def set_cell_margins(cell, **kwargs):
    """Set cell margins for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')

    for m_key, m_val in kwargs.items():
        m = OxmlElement(f'w:{m_key}')
        m.set(qn('w:w'), str(m_val))
        m.set(qn('w:type'), 'dxa')
        tcMar.append(m)

    tcPr.append(tcMar)

def set_table_cell_margins(table, top=None, bottom=None, start=None, end=None):
  """Set cell margins for all cells in a table."""
  for row in table.rows:
    for cell in row.cells:
      set_cell_margins(cell, top=top, bottom=bottom, start=start, end=end)

def insertHR(paragraph):
    p = paragraph._p  # p is the <w:p> XML element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pPr.insert_element_before(pBdr,
        'w:shd', 'w:tabs', 'w:suppressAutoHyphens', 'w:kinsoku', 'w:wordWrap',
        'w:overflowPunct', 'w:topLinePunct', 'w:autoSpaceDE', 'w:autoSpaceDN',
        'w:bidi', 'w:adjustRightInd', 'w:snapToGrid', 'w:spacing', 'w:ind',
        'w:contextualSpacing', 'w:mirrorIndents', 'w:suppressOverlap', 'w:jc',
        'w:textDirection', 'w:textAlignment', 'w:textboxTightWrap',
        'w:outlineLvl', 'w:divId', 'w:cnfStyle', 'w:rPr', 'w:sectPr',
        'w:pPrChange', 'w:header'
    )
    bottom = OxmlElement('w:top')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), 'auto')
    pBdr.append(bottom)

def submit_details():
    # Create a Word document for the resume
    document = Document()

    # Set default document font
    style = document.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'

    # Add Name
    name_heading = None
    if middle_name_entry.get() != "":
        name_heading = document.add_heading(first_name_entry.get().upper() + " " + middle_name_entry.get().upper() + " " + last_name_entry.get().upper(), level=0)
    else:
        name_heading = document.add_heading(first_name_entry.get().upper() + " " + last_name_entry.get().upper(), level=0)
    name_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = name_heading.runs[0]
    run.bold = True

    # Add Contact Information
    contact_info = f"{email_entry.get()} | {location_entry.get()} | {mobile_entry.get()}"
    contact_info_paragraph = document.add_paragraph(contact_info, style="Normal")
    if linkedin_entry.get() != "":
        contact_info_paragraph.add_run( " | ")
        add_hyperlink(contact_info_paragraph, "LinkedIn", linkedin_entry.get())
    if portfolio_entry.get() != "":
        contact_info_paragraph.add_run( " | ")
        add_hyperlink(contact_info_paragraph, "Portfolio", portfolio_entry.get())
    contact_info_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Add Summary
    summary_heading = document.add_heading("SUMMARY", level=2)
    run = summary_heading.runs[0]
    run.bold = True
    summary_paragraph = document.add_paragraph(summary_entry.get("1.0", "end").strip())
    insertHR(summary_paragraph)

    # Add Work Experience
    work_experience_heading = document.add_heading("WORK EXPERIENCE", level=2)
    run = work_experience_heading.runs[0]
    run.bold = True
    temp_paragraph = document.add_paragraph("")
    insertHR(temp_paragraph)
    for i in range(num_experiences):
        work_experience_table = document.add_table(2, 2)
        # Set cell margins (e.g., 50 twentieths of a point)
        set_table_cell_margins(work_experience_table, top=0, bottom=0)

        role = work_role_entries[i].get()
        company = work_company_entries[i].get()
        years = work_dates_entries[i].get()
        location = work_location_entries[i].get()

        run = work_experience_table.cell(i*2,0).paragraphs[0].add_run(f"{company.upper()}")
        run.bold = True
        run = work_experience_table.cell(i*2,1).paragraphs[0].add_run(f"{location}")
        run.bold = True
        work_experience_table.cell(i*2,1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = work_experience_table.cell(i*2+1,0).paragraphs[0].add_run(f"{role}")
        run.bold = True
        run.italic = True
        run = work_experience_table.cell(i*2+1,1).paragraphs[0].add_run(f"{years}")
        run.bold = True
        work_experience_table.cell(i*2+1,1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
        document.add_paragraph(f"Explain what you did, how you did it, mention impact with numbers and metrics if available, mention project.", style="List Bullet")
        document.add_paragraph(f"If it’s a present role, all verbs must be in a present continuous tense (e.g. establishing, maintaining). If it’s a past role, all verbs must be in the past tense (e.g. established, maintained).", style="List Bullet")
        document.add_paragraph(f"If you held a remote role, be sure to mention it at the corner of the country or location, like this: (Remote).", style="List Bullet")
        document.add_paragraph(f"E.g. Drove redevelopment of internal tracking system in use by 125 employees, resulting in 20+ new features, reduction of 20% in save/load time and 15% operation time", style="List Bullet")

    # Add Education
    education_heading = document.add_heading("EDUCATION", level=2)
    run = education_heading.runs[0]
    run.bold = True
    temp_paragraph = document.add_paragraph("")
    insertHR(temp_paragraph)
    for i in range(num_education):
        education_table = document.add_table(2, 2)
        # Set cell margins (e.g., 50 twentieths of a point)
        set_table_cell_margins(education_table, top=0, bottom=0)
        degree = edu_degree_entries[i].get()
        major = edu_major_entries[i].get()
        institution = edu_institution_entries[i].get()
        years = edu_years_entries[i].get()
        location = edu_location_entries[i].get()

        run = education_table.cell(i*2,0).paragraphs[0].add_run(f"{institution.upper()}")
        run.bold = True
        run = education_table.cell(i*2,1).paragraphs[0].add_run(f"{location}")
        run.bold = True
        education_table.cell(i*2,1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = education_table.cell(i*2+1,0).paragraphs[0].add_run(f"{degree} in {major}")
        run.bold = True
        run.italic = True
        run = education_table.cell(i*2+1,1).paragraphs[0].add_run(f"{years}")
        run.bold = True
        education_table.cell(i*2+1,1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
        document.add_paragraph(f"Awards: HERE", style="List Bullet")
        document.add_paragraph(f"Grade: HERE", style="List Bullet")
        document.add_paragraph(f"Project: HERE", style="List Bullet")

    # Add Additional Information
    additional_information_heading = document.add_heading("ADDITIONAL INFORMATION", level=2)
    run = additional_information_heading.runs[0]
    run.bold = True
    first_paragraph = document.add_paragraph(f"Technical Skills: HERE (mention technologies you are proficient in, e.g. Microsoft, Canva, Trello, HTML,  copy from job description, etc)", style="List Bullet")
    insertHR(first_paragraph)
    document.add_paragraph(f"Languages: HERE e.g. Spanish (LEVEL e.g. Native)", style="List Bullet")
    document.add_paragraph(f"Certifications: HERE (Courses, Internship .e.g. CFA Level 2 (August 2016))", style="List Bullet")
    document.add_paragraph(f"Awards: HERE (Educational, Career e.g.  Forbes Top 30 Entrepreneur 2023)", style="List Bullet")

    # Add Note
    note_heading = document.add_heading("NOTE:", level=2)
    run = note_heading.runs[0]
    run.bold = True
    document.add_paragraph(f"Ensure to save your final draft in .pdf format with your name like this “John_Doe” not “updated cv”!", style="List Bullet")
    document.add_paragraph(f"Ensure your email address, links, and resume details are all well spelt and punctuated because, honestly, nobody wants to give the recruiter a headache (they’ve seen enough already, lol). ", style="List Bullet")
    document.add_paragraph(f"While we can’t promise this template will land you a job overnight, we can guarantee it’ll give you a serious edge and have you standing out in the best way possible!", style="List Bullet")
    contact_us_paragraph = document.add_paragraph(f"To get a resume review, ", style="List Bullet")
    add_hyperlink(contact_us_paragraph, "CONTACT US", "https://wefind.space/contact/")

    #Add Skills
    # document.add_heading("Skills", level=2)
    # skills = skills_entry.get("1.0").split(",")  # Assume skills are comma-separated
    # first_skill = True
    # for skill in skills:
    #     skill_paragraph = document.add_paragraph(f"{skill.strip()}", style="List Bullet")
    #     if first_skill:
    #         insertHR(skill_paragraph)
    #         first_skill = False

    # Save the document
    file_path = filedialog.asksaveasfilename(defaultextension=".docx",
                                                filetypes=[("Word Documents", "*.docx")],
                                                title="Save Resume As")
    if file_path:
        document.save(file_path)


# Initialize the tkinter app
app = tk.Tk()
app.title("Resume Maker")

# Personal Information
personal_information_label = tk.Label(app, text="Personal Information:", font=("Helvetica", 12, "bold")).grid(row=0, column=0, sticky="w", pady=5)
personal_information_entry = tk.Entry(app)

tk.Label(app, text="First Name:").grid(row=1, column=0, sticky="w", padx=5)
first_name_entry = tk.Entry(app)
first_name_entry.grid(row=1, column=1)

tk.Label(app, text="Middle Name:").grid(row=2, column=0, sticky="w", padx=5)
middle_name_entry = tk.Entry(app)
middle_name_entry.grid(row=2, column=1)

tk.Label(app, text="Last Name:").grid(row=3, column=0, sticky="w", padx=5)
last_name_entry = tk.Entry(app)
last_name_entry.grid(row=3, column=1)

tk.Label(app, text="Email Address:").grid(row=4, column=0, sticky="w", padx=5)
email_entry = tk.Entry(app)
email_entry.grid(row=4, column=1)

tk.Label(app, text="Location (City, Country):").grid(row=5, column=0, sticky="w", padx=5)
location_entry = tk.Entry(app)
location_entry.grid(row=5, column=1)

tk.Label(app, text="Mobile Number (+x xxx xxx xxxx):").grid(row=6, column=0, sticky="w", padx=5)
mobile_entry = tk.Entry(app)
mobile_entry.grid(row=6, column=1)

tk.Label(app, text="LinkedIn Profile:").grid(row=7, column=0, sticky="w", padx=5)
linkedin_entry = tk.Entry(app)
linkedin_entry.grid(row=7, column=1)

tk.Label(app, text="Portfolio Link:").grid(row=8, column=0, sticky="w", padx=5)
portfolio_entry = tk.Entry(app)
portfolio_entry.grid(row=8, column=1)

# Summary Section
summary_label = tk.Label(app, text="Summary:", font=("Helvetica", 12, "bold")).grid(row=9, column=0, sticky="w", pady=5)
summary_entry = tk.Text(app, height=4, width=50)
summary_entry.grid(row=10, column=0, columnspan=2)

# Skills Section
skills_label = tk.Label(app, text="Skills:", font=("Helvetica", 12, "bold")).grid(row=11, column=0, sticky="w", pady=5)
skills_entry = tk.Text(app, height=1, width=50)
skills_entry.grid(row=12, column=0, columnspan=2)

# Work Experience & Education Section
tk.Label(app, text="Work Experience & Education:", font=("Helvetica", 12, "bold")).grid(row=13, column=0, sticky="w", pady=5)

# Number of Entries
entry_label = tk.Label(app, text="Number of Work Experience Entries:").grid(row=14, column=0, sticky="w", padx=5)
num_experiences_entry = tk.Entry(app)
num_experiences_entry.grid(row=14, column=1)

education_label = tk.Label(app, text="Number of Education Entries:").grid(row=15, column=0, sticky="w", padx=5)
num_education_entry = tk.Entry(app)
num_education_entry.grid(row=15, column=1)

# Generate Fields Button
generate_button = tk.Button(app, text="Generate Fields", command=generate_fields)
generate_button.grid(row=16, column=0, columnspan=2,sticky="w")

# Submit Button
submit_button = tk.Button(app, text="Submit", command=submit_details)

# Start the application
app.mainloop()
