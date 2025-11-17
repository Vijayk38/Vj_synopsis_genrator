import os
import re
import io # NEW IMPORT: Used to handle the Word document in memory
import streamlit as st # NEW IMPORT: The core library for the web app

# Third-party libraries
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import google.generativeai as genai

# --- Gemini Generation Function (Slightly Modified for Streamlit Secrets) ---

@st.cache_resource
def get_gemini_model():
    """Initializes and returns the Gemini model."""
    try:
        # Access the API key from Streamlit's secrets
        api_key = st.secrets["GOOGLE_API_KEY"]
        genai.configure(api_key=api_key)
        return genai.GenerativeModel("gemini-2.5-flash")
    except KeyError:
        st.error("Error: GOOGLE_API_KEY not found in Streamlit secrets.")
        st.stop()
    except Exception as e:
        st.error(f"Error configuring Gemini: {e}")
        st.stop()


def generate_report_with_gemini(topic, model):
    """Generates the report content using the Gemini model."""
    
    prompt = f"""
    Generate a **comprehensive, detailed, and professionally structured project report** for the topic: {topic}. 
    
    CRITICAL TONE AND STYLE INSTRUCTIONS:
    1. Write in an **active voice**, demonstrating critical analysis and nuanced understanding.
    2. Use sophisticated academic language and strong transition words to ensure the text flows naturally, avoiding a robotic or repetitive style.
    3. Back up claims with logical arguments, making the information feel authoritative and expert-written.
    
    To reach the requirement of 5 to 7 pages, the total text must be between **3000 and 4500 words**.
    
    Structure the report with two levels of headings, ensuring clear numbering followed by a colon for parsing:
    
    1. Introduction: (Define the problem, background, and scope)
    1.1 Background and Context:
    1.2 Problem Statement:
    
    2. Literature Review: (Analyze existing solutions and research gaps)
    2.1 Current State of Research:
    2.2 Identification of Gaps:
    
    3. Project Objectives: (Clear, measurable goals)
    3.1 Specific Aims:
    3.2 Deliverables:
    
    4. Detailed Methodology and System Design: (Technical approach, materials, and procedures)
    4.1 Proposed Architecture:
    4.2 Data Collection and Analysis Methods:
    
    5. Implementation Details and Results Analysis: 
    5.1 Execution Plan:
    5.2 Expected Outcomes and Validation:
    
    6. Conclusion and Future Work: 
    6.1 Summary of Findings:
    6.2 Future Scope and Recommendations:

    Format the content with numbered section headers followed by a colon (e.g., '1. Introduction:', '1.1 Background and Context:', etc.). 
    Ensure a blank line separates each section header and its preceding paragraph.
    Do NOT use any Markdown (like #, *, or **).
    """
    
    try:
        response = model.generate_content(prompt, generation_config={"temperature": 0.9}) 
        return response.text, None
    except Exception as e:
        return None, f"Report Generation Error: {str(e)}. Please check your API key and input."

# --- Word Conversion Function (MODIFIED to use io.BytesIO) ---

def text_to_word_buffer(text_content, topic):
    """
    Parses the text content and formats it into a Word document.
    Returns a bytes buffer (io.BytesIO) instead of saving to a file.
    """
    try:
        document = Document()
        
        # Section margins and styles (same as before)
        section = document.sections[0]
        section.top_margin = Inches(1); section.bottom_margin = Inches(1)
        section.left_margin = Inches(1); section.right_margin = Inches(1)

        topic_heading = document.add_heading(topic.upper(), level=0)
        topic_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        topic_heading.runs[0].font.size = Pt(20)
        topic_heading.runs[0].font.name = 'Times New Roman'
        
        document.add_paragraph(); document.add_paragraph() 
        
        lines = text_content.strip().split('\n')
        
        for line in lines:
            clean_line = line.strip()
            if not clean_line:
                document.add_paragraph()
                continue
            
            header_match = re.match(r'^(\d+(\.\d+)*)\s*(.*?):$', clean_line)

            if header_match:
                numbering = header_match.group(1) 
                title = header_match.group(3).strip().upper()

                # Level 1 headings (e.g., '1.')
                if numbering.count('.') == 0:
                    heading = document.add_heading(f"{numbering} {title}", level=1)
                    heading.runs[0].font.size = Pt(16)
                    heading.runs[0].font.name = 'Times New Roman'
                # Level 2 headings (e.g., '1.1')
                else:
                    heading = document.add_heading(f"{numbering} {title}", level=2)
                    heading.runs[0].font.size = Pt(14)
                    heading.runs[0].font.name = 'Times New Roman'
            else:
                # Normal paragraph text
                paragraph = document.add_paragraph(clean_line)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
                paragraph.runs[0].font.name = 'Times New Roman'
                paragraph.runs[0].font.size = Pt(12)
                
                paragraph_format = paragraph.paragraph_format
                paragraph_format.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                paragraph_format.line_spacing = 1.5 
        
        # Save the document to an in-memory byte buffer
        doc_io = io.BytesIO()
        document.save(doc_io)
        doc_io.seek(0) # Rewind the buffer to the beginning
        
        return doc_io, None
    except Exception as e:
        return None, f"Word Conversion Error: {str(e)}"

# --- Streamlit Main App ---

def main():
    st.set_page_config(
        page_title="Gemini AI Project Report Generator",
        layout="centered"
    )

    st.title("📄 AI-Powered Project Report Generator")
    st.markdown("Enter a project topic below. Gemini will generate a professional, structured, and in-depth **5-7 page** (3000-4500 words) report that you can download as a DOCX file.")
    st.markdown("---")

    # 1. Input Field (Replaces tkinter.Entry)
    topic = st.text_input(
        "Enter Your Project Topic:", 
        placeholder="e.g., The Impact of Quantum Computing on Financial Cryptography"
    )

    # 2. Submit Button (Replaces tkinter.Button)
    if st.button("Generate & Download Report", type="primary"):
        
        if not topic:
            st.error("Please enter a project topic to begin generation.")
            return

        # Initialize the model once
        model = get_gemini_model()
        
        st.info("Generation started. This process is intensive and may take up to **60 seconds** to complete. Please wait...")

        # --- Content Generation ---
        with st.spinner("Step 1/2: Generating the 3000+ word report content..."):
            report_text, error = generate_report_with_gemini(topic, model)

        if error:
            st.error(error)
            return

        # --- Word Conversion ---
        with st.spinner("Step 2/2: Formatting content into a professional DOCX document..."):
            doc_buffer, word_error = text_to_word_buffer(report_text, topic)

        if word_error:
            st.error(word_error)
            return

        # --- Download Button (The Streamlit Magic) ---
        
        st.success("✅ Report Generation Complete! Download your DOCX file below.")
        
        filename = f"{topic.replace(' ', '_').replace('/', '_')}_report.docx"
        
        # Create a download button for the in-memory buffer
        st.download_button(
            label="Download DOCX Report",
            data=doc_buffer,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
        st.markdown(f"*(Words Generated: ~{len(report_text.split())} words)*")

if __name__ == "__main__":
    main()