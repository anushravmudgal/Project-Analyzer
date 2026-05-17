import streamlit as st
import docx
import re
import io

# ==========================================
# ANALYZER CLASS
# ==========================================

class ProjectFileAnalyzer:
    def __init__(self, file_bytes):
        """Initializes the analyzer and loads the Word document."""
        try:
            self.doc = docx.Document(io.BytesIO(file_bytes))
            self.full_text = [p.text for p in self.doc.paragraphs if p.text.strip()]
            for table in self.doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        text = cell.text.strip()
                        if text and text not in self.full_text:
                            self.full_text.append(text)
        except Exception as e:
            st.error(f"Error loading document: {e}")
            self.doc = None

    def check_front_pages(self, mentor_name, subject_code):
        """Scans the document front matter for required details.
        Scans up to the first Heading 1 or the first 150 blocks, whichever comes first.
        """
        if not self.doc:
            return {}

        front_blocks = []
        for p in self.doc.paragraphs:
            if p.style.name == "Heading 1" and front_blocks:
                break
            front_blocks.append(p.text)
            if len(front_blocks) >= 150:
                break

        front_text = " ".join(front_blocks).lower()

        checks = {
            f"Subject Code ({subject_code})": False,
            f"Mentor Name ({mentor_name})": False,
            "Institution (IMS Noida)": False,
            "University (Chaudhary Charan Singh University)": False,
            "Proforma for Approval": False,
            "Certificate of Originality": False,
        }

        if subject_code.lower() in front_text:
            checks[f"Subject Code ({subject_code})"] = True
        if mentor_name.lower() in front_text:
            checks[f"Mentor Name ({mentor_name})"] = True
        if "ims noida" in front_text or "institute of management studies" in front_text:
            checks["Institution (IMS Noida)"] = True
        if "chaudhary charan singh" in front_text or "ccsu" in front_text:
            checks["University (Chaudhary Charan Singh University)"] = True
        if "proforma" in front_text:
            checks["Proforma for Approval"] = True
        if "certificate of originality" in front_text:
            checks["Certificate of Originality"] = True

        return checks

    def verify_index_table(self):
        """Strictly verifies if an Index table matches the benchmark format and content.
        BUG FIX: Changed table.rows.cells → table.rows[0].cells
        """
        if not self.doc:
            return False, "Document not loaded."

        index_table = None

        for table in self.doc.tables:
            if len(table.rows) > 0:
                # FIX: was table.rows.cells (AttributeError) → now table.rows[0].cells
                first_row_text = [cell.text.strip().lower() for cell in table.rows[0].cells]
                if any("description" in text for text in first_row_text) and any(
                    "page" in text for text in first_row_text
                ):
                    index_table = table
                    break

        if not index_table:
            return (
                False,
                "No structured Index table found. You must use a table with 'Description' and 'Page No.' columns.",
            )

        table_text = []
        for row in index_table.rows:
            for cell in row.cells:
                table_text.append(cell.text.strip().lower())

        full_index_text = " ".join(table_text)

        benchmark_chapters = [
            "about project",
            "system analysis",
            "system design",
            "coding",
            "testing",
            "system security",
            "cost estimation",
            "reports",
            "future scope",
            "limitations",
            "bibliography",
        ]

        missing_chapters = []
        for chapter in benchmark_chapters:
            if chapter not in full_index_text:
                missing_chapters.append(chapter.title())

        if missing_chapters:
            missing_str = ", ".join(missing_chapters)
            return False, f"Index table found, but missing mandatory chapters: {missing_str}."

        return True, "Index table matches the standard format and contains all mandatory chapters!"

    def check_mandatory_sections(self):
        """Checks if the main body contains standard project sections.
        Improved: matches against Heading-style paragraphs only to avoid false positives.
        """
        if not self.doc:
            return {}

        required_sections = [
            "Abstract",
            "Introduction",
            "Literature Review",
            "Methodology",
            "Results",
            "Conclusion",
            "References",
        ]
        found_sections = {section: False for section in required_sections}

        for paragraph in self.doc.paragraphs:
            # Only match headings to avoid false positives in body text
            is_heading = paragraph.style.name.startswith("Heading")
            text = paragraph.text.strip().lower()
            for section in required_sections:
                if section.lower() in text:
                    # Accept if it's a heading OR if it's a standalone line (bold/short)
                    if is_heading or len(text) < 50:
                        found_sections[section] = True

        return found_sections

    def analyze_structure_and_headings(self):
        """Extracts text formatted with Word's built-in Heading styles."""
        if not self.doc:
            return []
        headings = []
        for paragraph in self.doc.paragraphs:
            if paragraph.style.name.startswith("Heading"):
                headings.append(
                    {"level": paragraph.style.name, "text": paragraph.text.strip()}
                )
        return headings

    def check_abstract_length(self, min_words=150, max_words=300):
        """Finds the Abstract and checks its word count.
        Improved: uses .lower().startswith('abstract') for case-insensitive match.
        """
        if not self.doc:
            return None

        abstract_text = ""
        is_abstract = False

        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            # FIX: was exact match "abstract" → now case-insensitive startswith
            if text.lower().startswith("abstract") and not is_abstract:
                is_abstract = True
                continue
            if is_abstract and paragraph.style.name.startswith("Heading"):
                break
            if is_abstract:
                abstract_text += text + " "

        word_count = len(abstract_text.split())
        if word_count == 0:
            return "Abstract not found or empty.", "error"
        elif min_words <= word_count <= max_words:
            return f"Pass ({word_count} words)", "success"
        else:
            return f"Fail ({word_count} words). Should be {min_words}–{max_words}.", "warning"


# ==========================================
# SUBJECT CODE OPTIONS
# ==========================================

SUBJECT_CODES = [
    "BCA-605",
    "BCA-505P",
    "Custom — enter below",
]

# ==========================================
# STREAMLIT USER INTERFACE
# ==========================================

st.set_page_config(
    page_title="Project Report Analyzer", page_icon="📄", layout="centered"
)

st.title("📄 Major Project Report Analyzer")
st.write(
    "Ensure your project file meets the formatting, structural, and institutional "
    "guidelines before submission."
)

# --- Student Input Fields ---
st.info("Please provide your specific project details before uploading your file.")

col_input1, col_input2 = st.columns(2)

with col_input1:
    mentor_name = st.text_input(
        "Mentor's Name", placeholder="e.g., Mr. Anushrav Mudgal"
    )

with col_input2:
    selected_code = st.selectbox("Course & Subject Code", SUBJECT_CODES)
    if selected_code == "Custom — enter below":
        subject_code = st.text_input(
            "Enter custom subject code", placeholder="e.g., BBA-705P"
        )
    else:
        subject_code = selected_code
        st.caption(f"Selected: **{subject_code}**")

uploaded_file = st.file_uploader(
    "Upload your Word Document (.docx)", type=["docx"]
)

if uploaded_file is not None:
    if not mentor_name or not subject_code:
        st.warning(
            "⚠️ Please enter both your Mentor's Name and Subject Code above to begin the analysis."
        )
    else:
        file_bytes = uploaded_file.read()

        with st.spinner("Analyzing document structure and content..."):
            analyzer = ProjectFileAnalyzer(file_bytes)

            if analyzer.doc:
                st.divider()

                # 1. Front Page & Certificate Checks
                st.header("1. Title Page & Certificate Checks")
                st.write("Verifying standard institutional formats...")
                front_page_results = analyzer.check_front_pages(
                    mentor_name, subject_code
                )

                for detail, is_present in front_page_results.items():
                    if is_present:
                        st.success(f"✅ Found: {detail}")
                    else:
                        st.error(f"❌ Missing: {detail}")

                st.divider()

                # 2. Index Verification
                st.header("2. Index Verification")
                st.write("Verifying Index table format and mandatory chapters...")

                index_passed, index_msg = analyzer.verify_index_table()

                if index_passed:
                    st.success(f"✅ {index_msg}")
                    st.info(
                        "💡 *Note: Ensure your page numbers match manually, "
                        "as auto-checkers cannot verify dynamic Word page rendering.*"
                    )
                else:
                    st.error(f"❌ {index_msg}")
                    with st.expander("View Required Index Format"):
                        st.write(
                            "Your Index must be a Table containing exactly these major sections:"
                        )
                        st.markdown(
                            """
                        * About Project
                        * System Analysis
                        * System Design
                        * Coding
                        * Testing
                        * System Security Measures
                        * Cost Estimation of the Project
                        * Reports
                        * Future Scope and Further Enhancement
                        * Limitations of the Project
                        * Bibliography
                        """
                        )

                st.divider()

                # 3. General Word Count
                st.header("3. General Word Count")
                total_words = sum(len(text.split()) for text in analyzer.full_text)
                st.info(f"**Total Word Count:** {total_words} words")

                # 4. Mandatory Sections
                st.header("4. Mandatory Text Sections")
                sections = analyzer.check_mandatory_sections()

                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Required Section**")
                with col2:
                    st.write("**Status**")

                for section, is_present in sections.items():
                    colA, colB = st.columns(2)
                    colA.write(section)
                    if is_present:
                        colB.success("✅ Found")
                    else:
                        colB.error("❌ Missing")

                # 5. Abstract Length
                st.header("5. Content Rules")
                result = analyzer.check_abstract_length()
                if result:
                    msg, status = result
                    if status == "success":
                        st.success(f"**Abstract Length Check:** {msg}")
                    elif status == "warning":
                        st.warning(f"**Abstract Length Check:** {msg}")
                    else:
                        st.error(f"**Abstract Length Check:** {msg}")

                # 6. Document Structure (Headings)
                st.header("6. Document Structure (Headings)")
                st.write(
                    "If sections are missing here, you did not use Word's official 'Heading' styles."
                )
                headings = analyzer.analyze_structure_and_headings()

                if not headings:
                    st.error(
                        "❌ No formal headings found. Please use Word's Heading styles "
                        "(Heading 1, Heading 2, etc.) instead of just making text bold."
                    )
                else:
                    with st.expander("View Extracted Headings"):
                        for h in headings:
                            level_match = re.search(r"\d+", h["level"])
                            indent_level = int(level_match.group()) if level_match else 1
                            indent = "&nbsp;" * (indent_level * 4)
                            st.markdown(
                                f"{indent} - **[{h['level']}]** {h['text']}",
                                unsafe_allow_html=True,
                            )

# ==========================================
# FOOTER
# ==========================================

st.divider()
st.markdown(
    """
    <div style="text-align: center; color: #888; font-size: 0.85rem; padding: 8px 0 4px;">
        Built by
        <a href="https://www.linkedin.com/in/anushrav-mudgal/" target="_blank"
           style="color: #0A66C2; text-decoration: none; font-weight: 600;">
            Anushrav Mudgal
        </a>
        &nbsp;·&nbsp;
        <a href="https://www.linkedin.com/in/anushrav-mudgal/" target="_blank"
           style="color: #0A66C2; text-decoration: none;">
            🔗 LinkedIn
        </a>
    </div>
    """,
    unsafe_allow_html=True,
)
