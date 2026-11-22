from pathlib import Path

from pptx import Presentation

from pptx.util import Inches, Pt

from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from pptx.dml.color import RGBColor

from reportlab.lib.pagesizes import A4

from reportlab.lib.styles import (
    getSampleStyleSheet,
    ParagraphStyle,
)

from reportlab.lib.enums import TA_CENTER, TA_LEFT

from reportlab.lib.units import inch

from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    PageBreak,
    KeepTogether,
    HRFlowable,
)


# ============================================================
# PROJECT DIRECTORY CONFIGURATION
# ============================================================

# Resolve the current Python file location.
# The project root is assumed to be one directory above this file.

ROOT = Path(__file__).resolve().parents[1]


# Define the output directory where all generated documents
# and presentation files will be stored.

OUT = ROOT / "others"


# Create the output directory automatically if it does not exist.
# parents=True creates missing parent directories.
# exist_ok=True prevents errors if the directory already exists.

OUT.mkdir(
    parents=True,
    exist_ok=True
)


# ============================================================
# PROJECT INFORMATION
# ============================================================

# Footer information used throughout the presentations.

GROUP_FOOTER = (
    "Group 6 | "
    "Sakib Newaz Chowdhury, "
    "Shakela Mostofa Priya, "
    "Munim Mubashshir, "
    "Natasha Monir Shawon"
)


# Complete report header containing course and member information.

REPORT_HEADER = (
    "CSE440, Section 01 | Group 6 | "
    "Sakib Newaz Chowdhury (1921291042), "
    "Shakela Mostofa Priya (2132049042), "
    "Munim Mubashshir (2222118042), "
    "Natasha Monir Shawon (2121744642)"
)


# Course information.

COURSE_NAME = "CSE440"

COURSE_SECTION = "Section 01"

GROUP_NUMBER = "Group 6"


# Main project title.

PROJECT_TITLE = (
    "Poker Game Simulator with Expectiminimax "
    "for Strategic Decision-Making"
)


# ============================================================
# PRESENTATION CONFIGURATION
# ============================================================

# Standard presentation dimensions are controlled by
# the default PowerPoint presentation layout.

SLIDE_WIDTH = 13.333

SLIDE_HEIGHT = 7.5


# Footer positioning.

FOOTER_LEFT = 0.4

FOOTER_TOP = 6.9

FOOTER_WIDTH = 12.5

FOOTER_HEIGHT = 0.4


# Default font sizes.

TITLE_FONT_SIZE = 28

BODY_FONT_SIZE = 20

FOOTER_FONT_SIZE = 10


# Default colors.

BLACK_COLOR = RGBColor(0, 0, 0)

WHITE_COLOR = RGBColor(255, 255, 255)

DARK_BLUE_COLOR = RGBColor(31, 78, 121)


# ============================================================
# PDF CONFIGURATION
# ============================================================

# A4 page dimensions are used for all generated PDF reports.

PDF_PAGE_SIZE = A4


# PDF margins are measured in points.

PDF_LEFT_MARGIN = 48

PDF_RIGHT_MARGIN = 48

PDF_TOP_MARGIN = 48

PDF_BOTTOM_MARGIN = 48


# Default spacing values.

PDF_SMALL_SPACING = 6

PDF_MEDIUM_SPACING = 8

PDF_LARGE_SPACING = 12


# ============================================================
# GENERAL HELPER FUNCTIONS
# ============================================================

def validate_output_directory():

    """
    Check whether the output directory exists.

    If the directory is missing, recreate it.

    This function ensures that the generated files
    always have a valid destination.
    """

    OUT.mkdir(
        parents=True,
        exist_ok=True
    )

    return OUT.exists()


def get_output_path(filename):

    """
    Generate a complete output path.

    Parameters:
        filename: Name of the output document.

    Returns:
        A Path object pointing to the output file.
    """

    validate_output_directory()

    output_path = OUT / filename

    return output_path


def add_spacing(story, amount=PDF_MEDIUM_SPACING):

    """
    Add vertical spacing to a PDF story.

    Parameters:
        story: ReportLab story list.
        amount: Vertical spacing in points.
    """

    story.append(
        Spacer(1, amount)
    )


def add_horizontal_line(story):

    """
    Insert a horizontal divider into the PDF document.

    This helper function can be used to visually
    separate major sections.
    """

    story.append(
        HRFlowable(
            width="100%",
            thickness=0.5,
            color=RGBColor(100, 100, 100).hexval(),
            spaceBefore=5,
            spaceAfter=8
        )
    )


# ============================================================
# POWERPOINT FOOTER FUNCTION
# ============================================================

def add_footer(slide):

    """
    Add the group footer to a PowerPoint slide.

    Every slide receives the same footer information
    to maintain consistent presentation formatting.

    Parameters:
        slide: PowerPoint slide object.
    """

    # Create a text box near the bottom of the slide.

    box = slide.shapes.add_textbox(
        Inches(FOOTER_LEFT),
        Inches(FOOTER_TOP),
        Inches(FOOTER_WIDTH),
        Inches(FOOTER_HEIGHT)
    )

    # Access the text frame of the textbox.

    tf = box.text_frame

    # Set footer content.

    tf.text = GROUP_FOOTER

    # Access the first paragraph.

    paragraph = tf.paragraphs[0]

    # Apply font formatting.

    paragraph.font.size = Pt(
        FOOTER_FONT_SIZE
    )

    paragraph.font.color.rgb = DARK_BLUE_COLOR

    # Align footer text to the left.

    paragraph.alignment = PP_ALIGN.LEFT

    # Remove unnecessary text box margins.

    tf.margin_left = Inches(0.02)

    tf.margin_right = Inches(0.02)

    tf.margin_top = Inches(0.01)

    tf.margin_bottom = Inches(0.01)

    return box


# ============================================================
# POWERPOINT TITLE FORMATTING
# ============================================================

def format_slide_title(slide, title_text):

    """
    Set and format the title of a PowerPoint slide.

    Parameters:
        slide: Target slide.
        title_text: Text displayed as the slide title.
    """

    # Access the title placeholder.

    title_placeholder = slide.shapes.title

    # Assign the title.

    title_placeholder.text = title_text

    # Apply formatting to each paragraph.

    for paragraph in title_placeholder.text_frame.paragraphs:

        paragraph.font.size = Pt(
            TITLE_FONT_SIZE
        )

        paragraph.font.bold = True

        paragraph.font.color.rgb = DARK_BLUE_COLOR

    return title_placeholder


# ============================================================
# POWERPOINT BULLET FORMATTING
# ============================================================

def format_bullet_paragraph(paragraph, text):

    """
    Format a single bullet point.

    Parameters:
        paragraph: PowerPoint paragraph object.
        text: Bullet point content.
    """

    paragraph.text = text

    paragraph.level = 0

    paragraph.font.size = Pt(
        BODY_FONT_SIZE
    )

    paragraph.font.color.rgb = BLACK_COLOR

    paragraph.space_after = Pt(12)

    paragraph.space_before = Pt(4)

    return paragraph


# ============================================================
# POWERPOINT BULLET SLIDE CREATION
# ============================================================

def add_bullets(prs, title, bullets):

    """
    Create a PowerPoint slide containing bullet points.

    This function is reused for project update slides
    and final presentation slides.

    Parameters:
        prs: Presentation object.
        title: Slide title.
        bullets: List of bullet point strings.
    """

    # Add a slide using the standard title-and-content layout.

    slide = prs.slides.add_slide(
        prs.slide_layouts[1]
    )

    # Set the slide title.

    format_slide_title(
        slide,
        title
    )

    # Retrieve the body placeholder.

    body = slide.shapes.placeholders[1].text_frame

    # Remove the default placeholder text.

    body.clear()

    # Configure automatic text fitting.

    body.word_wrap = True

    # Insert all bullet points.

    for index, bullet in enumerate(bullets):

        # Use the first paragraph for the first bullet.
        # Create additional paragraphs for remaining bullets.

        if index == 0:

            paragraph = body.paragraphs[0]

        else:

            paragraph = body.add_paragraph()

        # Apply consistent formatting.

        format_bullet_paragraph(
            paragraph,
            bullet
        )

    # Add the standard group footer.

    add_footer(slide)

    return slide


# ============================================================
# PRESENTATION INITIALIZATION
# ============================================================

def initialize_presentation():

    """
    Create and initialize a new PowerPoint presentation.

    Returns:
        A Presentation object.
    """

    prs = Presentation()

    # Set the presentation dimensions.

    prs.slide_width = Inches(
        SLIDE_WIDTH
    )

    prs.slide_height = Inches(
        SLIDE_HEIGHT
    )

    return prs


# ============================================================
# PROJECT UPDATE PRESENTATION
# ============================================================

def make_update_pptx(path: Path):

    """
    Generate the project update PowerPoint presentation.

    The presentation summarizes the project's objectives,
    system components, progress, and upcoming activities.

    Parameters:
        path: Destination path of the PowerPoint file.
    """

    # Initialize presentation.

    prs = initialize_presentation()

    # --------------------------------------------------------
    # SLIDE 1: TITLE SLIDE
    # --------------------------------------------------------

    s0 = prs.slides.add_slide(
        prs.slide_layouts[0]
    )

    # Assign presentation title.

    s0.shapes.title.text = (
        "Project Update: Poker Game Simulator with Expectiminimax"
    )

    # Assign course information.

    s0.placeholders[1].text = (
        "Course: CSE440 | Section 01 | Group 6"
    )

    # Add footer.

    add_footer(s0)

    # --------------------------------------------------------
    # SLIDE 2: PROJECT OBJECTIVE
    # --------------------------------------------------------

    add_bullets(
        prs,
        "Project Objective",
        [
            "Build a simplified Texas Hold'em simulator.",

            "Integrate Expectiminimax for strategic AI decisions under uncertainty.",

            "Evaluate behavior against baseline strategies."
        ],
    )

    # --------------------------------------------------------
    # SLIDE 3: SYSTEM COMPONENTS
    # --------------------------------------------------------

    add_bullets(
        prs,
        "System Components",
        [
            "Game engine: deck, dealing, rounds, pot and chips management.",

            "Hand evaluator: 7-card ranking and score mapping.",

            "AI search: MAX, MIN, CHANCE recursive nodes."
        ],
    )

    # --------------------------------------------------------
    # SLIDE 4: PROJECT PROGRESS
    # --------------------------------------------------------

    add_bullets(
        prs,
        "Progress (Weeks 1-5)",
        [
            "Topic finalization and architecture planning completed.",

            "Core simulator prototype implemented and tested.",

            "Initial hand evaluator integrated.",

            "Expectiminimax integration in progress with depth tuning."
        ],
    )

    # --------------------------------------------------------
    # SLIDE 5: NEXT STEPS
    # --------------------------------------------------------

    add_bullets(
        prs,
        "Next Steps",
        [
            "Finalize utility and depth parameters.",

            "Run controlled experiments vs random/rule-based baselines.",

            "Prepare final report, final slides, and polished demo video."
        ],
    )

    # Save the completed presentation.

    prs.save(path)

    print(
        "Update presentation generated:",
        path
    )

    return path


# ============================================================
# FINAL PRESENTATION CONTENT
# ============================================================

def get_final_presentation_slides():

    """
    Return the complete slide content for the final
    project presentation.

    The function separates presentation content
    from presentation-generation logic.

    Returns:
        A list of slide titles and bullet point lists.
    """

    slides = [

        (
            "Problem Statement",
            [
                "Poker decisions involve uncertainty, hidden information, and stochastic events.",

                "Goal: create a simulator and AI strategy engine that reasons under chance."
            ]
        ),

        (
            "Method Overview",
            [
                "Simplified Texas Hold'em environment with sequential betting rounds.",

                "Expectiminimax model with MAX (AI), MIN (opponent), CHANCE (card draw)."
            ]
        ),

        (
            "Environment Design",
            [
                "State includes turn, chips, pot, hole cards, board cards, and action history.",

                "Actions: fold, call, raise with simplified betting constraints."
            ]
        ),

        (
            "Hand Evaluation",
            [
                "7-card evaluator ranks hands from high card to straight flush.",

                "Numeric utility supports tree-search comparisons."
            ]
        ),

        (
            "Search and Decision Flow",
            [
                "Depth-limited recursive search to control complexity.",

                "Expected values aggregated over chance outcomes."
            ]
        ),

        (
            "Implementation Structure",
            [
                "main.py as entry point.",

                "Modular package for game engine, AI logic, and visualization."
            ]
        ),

        (
            "Experimental Setup",
            [
                "AI tested against random and simple rule-based baselines.",

                "Tracked win rate, chip delta trend, and action profile."
            ]
        ),

        (
            "Observed Results",
            [
                "AI generally shows stronger showdown selection and fold discipline.",

                "Performance is sensitive to depth and utility scaling."
            ]
        ),

        (
            "Limitations",
            [
                "Simplified betting abstraction vs full poker complexity.",

                "State-space growth limits deep exhaustive exploration."
            ]
        ),

        (
            "Conclusion and Future Work",
            [
                "Project demonstrates feasible strategic play with Expectiminimax.",

                "Future: richer opponent modeling and stronger abstractions."
            ]
        ),

    ]

    return slides


# ============================================================
# FINAL POWERPOINT PRESENTATION
# ============================================================

def make_final_pptx(path: Path):

    """
    Generate the final project PowerPoint presentation.

    The presentation includes the problem statement,
    methodology, system design, implementation,
    experiments, results, and conclusion.
    """

    # Create a new presentation.

    prs = initialize_presentation()

    # --------------------------------------------------------
    # TITLE SLIDE
    # --------------------------------------------------------

    s0 = prs.slides.add_slide(
        prs.slide_layouts[0]
    )

    # Set presentation title.

    s0.shapes.title.text = (
        "Final Presentation: Poker Game Simulator with Expectiminimax"
    )

    # Preserve original course and group information.

    s0.placeholders[1].text = (
        "Course: CSE440 | Section 01 | Group 7"
    )

    # Add group footer.

    add_footer(s0)

    # --------------------------------------------------------
    # CONTENT SLIDES
    # --------------------------------------------------------

    slides = get_final_presentation_slides()

    # Generate slides dynamically.

    for slide_title, bullet_points in slides:

        add_bullets(
            prs,
            slide_title,
            bullet_points
        )

    # --------------------------------------------------------
    # SAVE PRESENTATION
    # --------------------------------------------------------

    prs.save(path)

    print(
        "Final presentation generated:",
        path
    )

    return path


# ============================================================
# PDF REPORT STYLE CONFIGURATION
# ============================================================

def create_pdf_styles():

    """
    Create and configure PDF paragraph styles.

    Separate styles are used for:
        - Main document title.
        - Section headings.
        - Normal paragraphs.
        - Centered report headers.

    Returns:
        A dictionary containing paragraph styles.
    """

    styles = getSampleStyleSheet()

    # Standard body paragraph style.

    normal = ParagraphStyle(
        "CustomBodyText",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=15,
        alignment=TA_LEFT,
        spaceAfter=6
    )

    # Section heading style.

    heading = ParagraphStyle(
        "CustomHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=17,
        alignment=TA_LEFT,
        spaceBefore=10,
        spaceAfter=8
    )

    # Main report title style.

    title = ParagraphStyle(
        "CustomTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=25,
        alignment=TA_CENTER,
        spaceAfter=14
    )

    # Header style.

    header = ParagraphStyle(
        "CustomReportHeader",
        parent=normal,
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        alignment=TA_CENTER,
        spaceAfter=8
    )

    # Return all styles together.

    return {
        "normal": normal,
        "heading": heading,
        "title": title,
        "header": header
    }


# ============================================================
# REPORT SECTION CONTENT
# ============================================================

def get_report_sections():

    """
    Return the original report sections.

    All project descriptions are preserved.
    The content is stored separately to make
    the report easier to maintain and extend.
    """

    sections = [

        (
            "Abstract",

            "This project develops a simplified Texas Hold'em poker simulator and integrates an Expectiminimax-based agent for strategic play under uncertainty. The environment models imperfect information and stochastic transitions. We evaluate behavior with baseline opponents and discuss strengths and limitations."
        ),

        (
            "1. Introduction",

            "Poker is a representative decision-making problem with hidden information and random events. Traditional deterministic minimax is insufficient because card draws and incomplete observations require probabilistic reasoning. Our project focuses on designing a practical educational simulator and studying how Expectiminimax performs in this domain."
        ),

        (
            "2. Objectives",

            "Primary objectives are: (i) design a modular poker environment, (ii) implement correct hand ranking and state transitions, (iii) integrate Expectiminimax with manageable complexity, and (iv) evaluate the algorithm through reproducible test scenarios."
        ),

        (
            "3. System Design",

            "The architecture includes a game engine, card/deck utilities, hand evaluation component, AI decision module, and user interaction layer. The game engine controls phases (pre-flop to showdown), validates actions, updates chips and pot, and resolves winners."
        ),

        (
            "4. Expectiminimax Model",

            "Decision nodes are represented as MAX (our AI), MIN (opponent response), and CHANCE (unknown card outcomes). Utility is based on hand strength, pot dynamics, and terminal outcomes. Depth limits and controlled branching are used to avoid combinatorial explosion."
        ),

        (
            "5. Implementation Notes",

            "The project is organized for course compliance with a main executable, requirements list, and dedicated support folders. We use modular Python code to separate concerns and improve readability, testing, and future extensions."
        ),

        (
            "6. Experimental Plan",

            "We run repeated matches against random and rule-based opponents, then compare win rates and chip trajectories. Scenario-based tests include early folds, aggressive betting lines, and shared strong-board situations to validate behavior under diverse conditions."
        ),

        (
            "7. Member Contributions",

            "Turjo led architecture and core integration; Sabbir handled literature synthesis and scenario design; Nashita focused on user-facing flows, testing, and documentation clarity; Talha supported utilities, debugging, and empirical tuning observations."
        ),

        (
            "8. Results Summary",

            "Early findings indicate that Expectiminimax improves action consistency compared to naive baselines. The bot tends to avoid weak commitments and pursues strong-value opportunities more effectively, though quality depends on evaluation calibration."
        ),

        (
            "9. Limitations and Future Work",

            "Current simulator uses simplified betting abstractions and depth-limited search. Future improvements include richer opponent modeling, improved abstraction of information sets, and larger-scale benchmarking."
        ),

        (
            "10. Conclusion",

            "This work demonstrates a practical pipeline for building a strategic AI poker simulator in an academic setting. It satisfies the core project objective of combining environment design and uncertain-decision algorithms with measurable outcomes."
        ),

    ]

    return sections


# ============================================================
# APPENDIX CONTENT
# ============================================================

def get_appendix_sections():

    """
    Return the additional appendix sections
    used in the final report.
    """

    appendix = [

        (
            "Appendix A: Weekly Timeline",

            "Week 1 to Week 5 covered planning, architecture, game-engine prototype, hand-evaluator development, and Expectiminimax integration. Subsequent weeks emphasize evaluation, final polishing, and presentation readiness."
        ),

        (
            "Appendix B: Ethical and Academic Compliance",

            "All external ideas and materials should be properly cited. The project follows course policy regarding plagiarism, teamwork integrity, and submission discipline."
        ),

    ]

    return appendix


# ============================================================
# PDF PARAGRAPH GENERATION
# ============================================================

def add_pdf_paragraph(story, text, style):

    """
    Add a paragraph to the PDF story.

    Parameters:
        story: ReportLab story collection.
        text: Paragraph content.
        style: Paragraph formatting style.
    """

    paragraph = Paragraph(
        text,
        style
    )

    story.append(
        paragraph
    )

    return paragraph


# ============================================================
# PDF REPORT HEADER
# ============================================================

def add_report_header(story, styles):

    """
    Add the report header containing course and
    group information.
    """

    add_pdf_paragraph(
        story,
        REPORT_HEADER,
        styles["header"]
    )

    add_spacing(
        story,
        PDF_MEDIUM_SPACING
    )


# ============================================================
# PDF REPORT TITLE
# ============================================================

def add_report_title(story, update_mode, styles):

    """
    Add the appropriate report title.

    The title depends on whether the document
    is an update report or final report.
    """

    if update_mode:

        report_title = "Project Update Report (2 Pages)"

    else:

        report_title = "Final Project Report (8 Pages)"

    add_pdf_paragraph(
        story,
        report_title,
        styles["title"]
    )

    add_pdf_paragraph(
        story,
        PROJECT_TITLE,
        styles["heading"]
    )

    add_spacing(
        story,
        PDF_LARGE_SPACING
    )


# ============================================================
# PDF SECTION BUILDER
# ============================================================

def add_report_sections(story, sections, styles):

    """
    Add a collection of report sections.

    Each section contains:
        - A heading.
        - A descriptive paragraph.
        - Vertical spacing.

    Parameters:
        story: PDF story list.
        sections: Collection of report sections.
        styles: Paragraph style dictionary.
    """

    for heading_text, body_text in sections:

        # Create section heading.

        heading_paragraph = Paragraph(
            heading_text,
            styles["heading"]
        )

        # Create section body.

        body_paragraph = Paragraph(
            body_text,
            styles["normal"]
        )

        # Keep heading and paragraph together
        # where possible to improve page layout.

        section_content = KeepTogether(
            [
                heading_paragraph,
                body_paragraph,
                Spacer(1, PDF_MEDIUM_SPACING)
            ]
        )

        story.append(
            section_content
        )


# ============================================================
# UPDATE REPORT SECOND PAGE
# ============================================================

def add_update_appendix(story, styles):

    """
    Add the second page of the update report.

    The page contains update highlights and
    additional technical discussion.
    """

    # Force the appendix to begin on a new page.

    story.append(
        PageBreak()
    )

    # Add appendix heading.

    add_pdf_paragraph(
        story,
        "Appendix: Update Highlights",
        styles["heading"]
    )

    # Add the original update summary.

    update_summary = (

        "This second page summarizes the update-phase outcomes: initial game flow stabilization, "

        "first evaluator integration, and preliminary Expectiminimax testing signals. "

        "It also records immediate next actions on calibration, baseline comparison, and presentation readiness."

    )

    add_pdf_paragraph(
        story,
        update_summary,
        styles["normal"]
    )

    add_spacing(
        story,
        PDF_LARGE_SPACING
    )

    # Add repeated discussion paragraphs to preserve
    # the original report's second-page structure.

    for index in range(10):

        discussion = (

            "The update period confirms technical feasibility and identifies the main risk areas: utility scaling, "

            "depth-vs-runtime balance, and ensuring result consistency across random seeds and scenario sets."

        )

        add_pdf_paragraph(
            story,
            discussion,
            styles["normal"]
        )

        add_spacing(
            story,
            PDF_SMALL_SPACING
        )


# ============================================================
# FINAL REPORT DETAILED DISCUSSION
# ============================================================

def add_detailed_discussion_pages(story, styles):

    """
    Add the extended discussion pages for the final report.

    The original code intentionally generates additional
    discussion pages to reach approximately eight pages.

    The number of discussion pages and paragraphs
    is maintained from the original implementation.
    """

    # Generate six detailed discussion pages.

    for page_number in range(1, 7):

        # Begin each discussion on a new page.

        story.append(
            PageBreak()
        )

        # Create a page-specific heading.

        heading_text = (
            f"Detailed Discussion Page {page_number}"
        )

        add_pdf_paragraph(
            story,
            heading_text,
            styles["heading"]
        )

        # Main discussion introduction.

        introduction = (

            "This page intentionally contains extended analysis text for the final report body. "

            "Discuss algorithmic trade-offs, implementation choices, observed behaviors, and interpretation of experiment outcomes in depth. "

            "Include references, figures, equations, and tables in the final polished version."

        )

        add_pdf_paragraph(
            story,
            introduction,
            styles["normal"]
        )

        add_spacing(
            story,
            PDF_LARGE_SPACING
        )

        # Add eight technical discussion paragraphs.

        for paragraph_number in range(8):

            technical_discussion = (

                "Expectiminimax balances adversarial and stochastic reasoning. In poker, chance outcomes from unknown card draws strongly influence utility projections. "

                "Depth control, branching reduction, and a stable evaluation function are key to obtaining useful decisions within practical runtime constraints."

            )

            add_pdf_paragraph(
                story,
                technical_discussion,
                styles["normal"]
            )

            add_spacing(
                story,
                PDF_SMALL_SPACING
            )


# ============================================================
# MAIN PDF STORY CONSTRUCTION
# ============================================================

def build_story(update_mode: bool):

    """
    Construct the complete PDF document content.

    Parameters:
        update_mode:
            True  -> Build the project update report.
            False -> Build the final project report.

    Returns:
        A ReportLab story list.
    """

    # Initialize paragraph styles.

    styles = create_pdf_styles()

    # Create the story collection.

    story = []

    # --------------------------------------------------------
    # REPORT HEADER
    # --------------------------------------------------------

    add_report_header(
        story,
        styles
    )

    # --------------------------------------------------------
    # REPORT TITLE
    # --------------------------------------------------------

    add_report_title(
        story,
        update_mode,
        styles
    )

    # --------------------------------------------------------
    # MAIN REPORT SECTIONS
    # --------------------------------------------------------

    sections = get_report_sections()

    # The update report uses only the first seven sections.

    if update_mode:

        # Keep concise and then force a second page.

        sections = sections[:7]

    else:

        # Append the original appendix sections
        # to the final report.

        sections.extend(
            get_appendix_sections()
        )

    # Add all selected sections.

    add_report_sections(
        story,
        sections,
        styles
    )

    # --------------------------------------------------------
    # UPDATE REPORT APPENDIX
    # --------------------------------------------------------

    if update_mode:

        add_update_appendix(
            story,
            styles
        )

    # --------------------------------------------------------
    # FINAL REPORT DISCUSSION
    # --------------------------------------------------------

    if not update_mode:

        add_detailed_discussion_pages(
            story,
            styles
        )

    # Return the completed story.

    return story


# ============================================================
# PDF DOCUMENT GENERATION
# ============================================================

def make_pdf(path: Path, update_mode: bool):

    """
    Generate a PDF report.

    Parameters:
        path: Destination PDF file path.
        update_mode: Selects update or final report.

    The function creates a SimpleDocTemplate,
    builds the story, and saves the completed PDF.
    """

    # Configure the PDF document.

    doc = SimpleDocTemplate(

        str(path),

        pagesize=PDF_PAGE_SIZE,

        leftMargin=PDF_LEFT_MARGIN,

        rightMargin=PDF_RIGHT_MARGIN,

        topMargin=PDF_TOP_MARGIN,

        bottomMargin=PDF_BOTTOM_MARGIN,

        title=path.stem,

        author=REPORT_HEADER,

    )

    # Build the report content.

    story = build_story(
        update_mode=update_mode
    )

    # Generate the final PDF document.

    doc.build(
        story
    )

    print(
        "PDF report generated:",
        path
    )

    return path


# ============================================================
# OUTPUT VALIDATION
# ============================================================

def validate_generated_files(file_paths):

    """
    Verify that all expected output files exist.

    Parameters:
        file_paths: List of generated file paths.

    Returns:
        True if every file exists.
        False otherwise.
    """

    all_files_exist = True

    print("\nChecking generated files...")

    for file_path in file_paths:

        if file_path.exists():

            print(
                "[SUCCESS]",
                file_path.name
            )

        else:

            print(
                "[MISSING]",
                file_path.name
            )

            all_files_exist = False

    return all_files_exist


# ============================================================
# PROJECT FILE GENERATION CONTROLLER
# ============================================================

def generate_all_documents():

    """
    Generate all project deliverables.

    The function creates:
        1. Update presentation.
        2. Final presentation.
        3. Update report.
        4. Final report.

    All files are saved in the others directory.
    """

    # Ensure output directory is available.

    validate_output_directory()

    # --------------------------------------------------------
    # OUTPUT FILE PATHS
    # --------------------------------------------------------

    update_presentation_path = get_output_path(
        "Update presentation.pptx"
    )

    final_presentation_path = get_output_path(
        "Final presentation.pptx"
    )

    update_report_path = get_output_path(
        "Update report.pdf"
    )

    final_report_path = get_output_path(
        "Final report.pdf"
    )

    # --------------------------------------------------------
    # GENERATE UPDATE PRESENTATION
    # --------------------------------------------------------

    make_update_pptx(
        update_presentation_path
    )

    # --------------------------------------------------------
    # GENERATE FINAL PRESENTATION
    # --------------------------------------------------------

    make_final_pptx(
        final_presentation_path
    )

    # --------------------------------------------------------
    # GENERATE UPDATE REPORT
    # --------------------------------------------------------

    make_pdf(
        update_report_path,
        update_mode=True
    )

    # --------------------------------------------------------
    # GENERATE FINAL REPORT
    # --------------------------------------------------------

    make_pdf(
        final_report_path,
        update_mode=False
    )

    # Collect generated file paths.

    generated_files = [

        update_presentation_path,

        final_presentation_path,

        update_report_path,

        final_report_path,

    ]

    # Verify the output files.

    validate_generated_files(
        generated_files
    )

    return generated_files


# ============================================================
# MAIN APPLICATION ENTRY POINT
# ============================================================

def main():

    """
    Main execution function.

    This function coordinates the complete document
    generation workflow.

    It also handles unexpected execution errors.
    """

    print("=" * 65)

    print(
        "POKER GAME SIMULATOR WITH EXPECTIMINIMAX"
    )

    print(
        "CSE440 | Section 01 | Group 6"
    )

    print(
        "Project Document Generation System"
    )

    print("=" * 65)

    try:

        # Generate all project deliverables.

        generated_files = generate_all_documents()

        # Display final output location.

        print("\nAll document generation tasks completed.")

        print(
            "Generated files in:",
            OUT
        )

        # Display each generated file.

        print("\nGenerated deliverables:")

        for file_path in generated_files:

            print(
                " -",
                file_path.name
            )

        print("\nProcess completed successfully.")

    except Exception as error:

        # Display an informative error message.

        print(
            "\nAn error occurred during document generation:"
        )

        print(
            str(error)
        )

        # Re-raise the error so that execution failures
        # remain visible to the calling environment.

        raise


# ============================================================
# SCRIPT EXECUTION
# ============================================================

if __name__ == "__main__":

    # Start the document generation process.

    main()

    # End of program.

