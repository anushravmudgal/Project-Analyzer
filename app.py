import streamlit as st
import docx
import re
import io

class ProjectFileAnalyzer:
    def __init__(self, file_bytes):
        """Initializes the analyzer and loads the Word document, extracting text from paragraphs and tables."""
        try:
            # Streamlit uploads files as bytes, so we wrap it in io.BytesIO
            self.doc = docx.Document(io.BytesIO(file_bytes))
            
            # Extract text from standard paragraphs
            self.full_text = [p.text for p in self.doc.paragraphs if p.text.strip()]
            
            # Extract text from tables (Crucial for reading the Index)
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
        """Scans the beginning of the document for required benchmark details."""
        if not self.doc: return {}
        
        # Combine the first 150 blocks of text to safely cover the first 4-5 pages
        front_text = " ".join(self.full_text[:150]).lower()
        
        checks = {
            f"Subject Code ({subject_code})": False,
            f"Mentor Name ({mentor_name})": False,
            "Institution (IMS Noida)": False,
            "University (Chaudhary Charan Singh University)": False,
            "Proforma for Approval": False,
            "Certificate of Originality": False
        }
        
        # Dynamic Checks based on user input
        if subject_code.lower() in front_text:
            checks[f"Subject Code ({subject_code})"] = True
        if mentor_name.lower() in front_text:
            checks[f"Mentor Name ({mentor_name})"] = True
            
        # Static Checks based on the standard format
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
        """Strictly verifies if an Index table matches the benchmark format and content."""
        if not self.doc: return False, "Document not loaded."
        
        index_table = None
        
        # 1. Look specifically for a Table with "Description" and "Page No" headers
        for table in self.doc.tables:
            if len(table.rows) > 0:
                first_row_text = [cell.text.strip().lower() for cell in table.rows.cells]
                if any('description' in text for text in first_row_text) and any('page' in text for text in first_row_text):
                    index_table = table
                    break
                    
        if not index_table:
            return False, "No structured Index table found. You must use a table with 'Description' and 'Page No.' columns."
            
        # 2. Extract all text from this specific table
        table_text = []
        for row in index_table.rows:
            for cell in row.cells:
                table_text.append(cell.text.strip().lower())
        
        full_index_text = " ".join(table_text)
        
        # 3. Benchmark Chapters based on the official standard
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
            "bibliography"
        ]
        
        missing_chapters = []
        for chapter in benchmark_chapters:
            if chapter not in full_index_text:
                missing_chapters.append(chapter.title())
                
        if missing_chapters:
            missing_str = ", ".join(missing_chapters)
            return False, f"Index table found, but it is missing mandatory chapters: {missing_str}."
            
        return True, "Index table matches the standard format and contains all mandatory chapters!"

    def check_mandatory_sections(self):
        """Checks if the main body of the document contains standard project sections."""
        if not self.doc: return {}
        required_sections = [
            "Abstract", "Introduction", "Literature Review", 
            "Methodology", "Results", "Conclusion", "References"
        ]
        found_sections = {section: False for section in required_sections}
        
        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip().lower()
            for section in required_sections:
                if section.lower() in text:
                    found_sections[section] = True
        return found_sections

    def analyze_structure_and_headings(self):
        """Extracts text formatted with Word's built-in Heading styles."""
        if not self.doc: return []
        headings = []
        for paragraph in self.doc.paragraphs:
            if paragraph.style.name.startswith('Heading'):
                headings.append({
                    "level": paragraph.style.name,
                    "text": paragraph.text.strip()
                })
        return headings

    def check_abstract_length(self, min_words=150, max_words=300):
        """Finds the Abstract and calculates its word count."""
        if not self.doc: return None
        abstract_text = ""
        is_abstract = False
        
        for paragraph in self.doc.paragraphs:
            text = paragraph.text.strip()
            if text.lower() == "abstract":
                is_abstract = True
                continue
            # Stop when the next major heading is reached
            if is_abstract and paragraph.style.name.startswith('Heading'):
                break
            if is_abstract:
                abstract_text += text + " "
                
        word_count = len(abstract_text.split())
        if word_count == 0:
            return "Abstract not found or empty.", "error"
        elif min_words <= word_count <= max_words:
            return f"Pass ({word_count} words)", "success"
        else:
            return f"Fail ({word_count} words). Should be {min_words}-{max_words}.", "warning"

# ==========================================
# STREAMLIT USER INTERFACE
# ==========================================
st.set_page_config(page_title="Project Report Analyzer", page_icon="📄", layout="centered")

st.title("📄 Major Project Report Analyzer")
st.write("Ensure your project file meets the formatting, structural, and institutional guidelines before submission.")

# --- Student Input Fields ---
st.info("Please provide your specific project details before uploading your file.")
col_input1, col_input2 = st.columns(2)
with col_input1:
    mentor_name = st.text_input("Mentor's Name", placeholder="e.g., Mr. Prateek Tiwari")
with col_input2:
    subject_code = st.text_input("Course & Subject Code", placeholder="e.g., BCA-605P")

uploaded_file = st.file_uploader("Upload your Word Document (.docx)", type=["docx"])

if uploaded_file is not None:
    if not mentor_name or not subject_code:
        st.warning("⚠️ Please enter both your Mentor's Name and Subject Code above to begin the analysis.")
    else:
        file_bytes = uploaded_file.read()
        
        with st.spinner("Analyzing document structure and content..."):
            analyzer = ProjectFileAnalyzer(file_bytes)
            
            if analyzer.doc:
                st.divider()
                
                # 1. Front Page Benchmark Checks
                st.header("1. Title Page & Certificate Checks")
                st.write("Verifying standard institutional formats...")
                front_page_results = analyzer.check_front_pages(mentor_name, subject_code)
                
                for detail, is_present in front_page_results.items():
                    if is_present:
                        st.success(f"✅ Found: {detail}")
                    else:
                        st.error(f"❌ Missing: {detail}")
                
                st.divider()
                
                # 2. Strict Index & Table of Contents Verification
                st.header("2. Index Verification")
                st.write("Verifying Index table format and mandatory chapters...")
                
                index_passed, index_msg = analyzer.verify_index_table()
                
                if index_passed:
                    st.success(f"✅ {index_msg}")
                    st.info("💡 *Note: Ensure your page numbers match manually, as auto-checkers cannot verify dynamic Word page rendering.*")
                else:
                    st.error(f"❌ {index_msg}")
                    with st.expander("View Required Index Format"):
                        st.write("Your Index must be a Table containing exactly these major sections:")
                        st.markdown("""
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
                        """)
                
                st.divider()
                
                # 3. General Overview
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
                msg, status = analyzer.check_abstract_length()
                if status == "success":
                    st.success(f"**Abstract Length Check:** {msg}")
                elif status == "warning":
                    st.warning(f"**Abstract Length Check:** {msg}")
                else:
                    st.error(f"**Abstract Length Check:** {msg}")

                # 6. Structural Headings
                st.header("6. Document Structure (Headings)")
                st.write("If sections are missing here, you did not use Word's official 'Heading' styles.")
                headings = analyzer.analyze_structure_and_headings()
                
                if not headings:
                    st.error("❌ No formal headings found. Please use Word's Heading styles (Heading 1, Heading 2, etc.) instead of just making text bold.")
                else:
                    with st.expander("View Extracted Headings"):
                        for h in headings:
                            level_match = re.search(r'\d+', h['level'])
                            indent_level = int(level_match.group()) if level_match else 1
                            indent = "&nbsp;" * (indent_level * 4)
                            st.markdown(f"{indent} - **[{h['level']}]** {h['text']}")