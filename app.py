
import streamlit as st
import os
import re
import json
import random
from typing import List, Dict, Tuple
import pypdf
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.units import inch

# ==========================================
# CONFIGURATION & SETUP
# ==========================================

st.set_page_config(
    page_title="Local AI Exam Generator (Offline)",
    page_icon="💻",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header { font-size: 2.5rem; font-weight: bold; color: #1f1f1f; text-align: center; margin-bottom: 1rem; }
    .sub-header { font-size: 1.2rem; color: #555; text-align: center; margin-bottom: 2rem; }
    .exam-paper { background-color: #ffffff; padding: 30px; border: 1px solid #ddd; box-shadow: 0 4px 6px rgba(0,0,0,0.1); border-radius: 5px; margin-top: 20px; font-family: 'Times New Roman', Times, serif; }
    .section-header { background-color: #f0f0f0; padding: 10px; font-weight: bold; margin-top: 20px; border-left: 5px solid #333; }
    .question-item { margin-bottom: 15px; padding: 10px 0; border-bottom: 1px dotted #ccc; }
    .download-box { border: 1px solid #eee; padding: 15px; border-radius: 10px; margin-top: 10px; background-color: #fafafa; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# SESSION STATE
# ==========================================

if 'step' not in st.session_state:
    st.session_state.step = 1
if 'full_text' not in st.session_state:
    st.session_state.full_text = ""
# We store chapters as a list of dicts: {'name': 'Title', 'content': 'Text...'}
if 'chapters_data' not in st.session_state:
    st.session_state.chapters_data = []
if 'generated_exam' not in st.session_state:
    st.session_state.generated_exam = None

# ==========================================
# LOCAL LOGIC ENGINE (NO API)
# ==========================================

def analyze_structure_locally(text: str) -> List[Dict]:
    """
    Analyzes text to find chapters or chunks without an API.
    It looks for patterns like 'Chapter 1', 'Unit 1', etc.
    If none found, it splits the text into equal chunks.
    """
    chapters = []
    
    # 1. Try to find explicit headings (Chapter, Unit, Section, Part)
    # Regex looks for lines that start with these keywords or follow a newline
    pattern = r'(?:\n|^)(Chapter|Unit|Section|Part)\s+\d+[:\-.]?\s*(.*)'
    matches = list(re.finditer(pattern, text, re.IGNORECASE))
    
    if len(matches) > 1:
        # Split based on found headings
        for i, match in enumerate(matches):
            start = match.start()
            # End is the start of the next match
            end = matches[i+1].start() if i < len(matches) - 1 else len(text)
            
            title = match.group(0).strip()
            content = text[start:end].strip()
            chapters.append({"name": title, "content": content})
    else:
        # 2. Fallback: Split by double newlines (paragraphs) to create "Sections"
        # or simply split into 5 equal chunks if no structure found
        st.info("No clear chapter headings found. Document is split into 5 equal sections automatically.")
        chunk_size = len(text) // 5
        for i in range(5):
            start = i * chunk_size
            end = (i + 1) * chunk_size if i < 4 else len(text)
            content = text[start:end].strip()
            if len(content) > 50: # Only keep non-empty chunks
                chapters.append({"name": f"Section {i+1}", "content": content})
                
    return chapters

def clean_sentence(s: str) -> str:
    """Remove non-alphanumeric chars except spaces for comparison."""
    return re.sub(r'[^a-zA-Z0-9\s]', '', s).strip()

def local_generate_exam(text_source: str, config: dict) -> Dict:
    """
    Generates questions using heuristics (No LLM).
    Strategy:
    1. Break text into sentences.
    2. MCQs: Take a sentence, mask a keyword (Cloze), create distractors.
    3. Short/Long: Select sentences and ask to explain.
    """
    # Extract sentences
    # Split by ., !, ? followed by space or newline
    raw_sentences = re.split(r'(?<=[.!?])\s+', text_source)
    
    # Filter: Keep sentences that are decent length (20+ chars)
    sentences = [s.strip() for s in raw_sentences if len(s.strip()) > 20 and len(s.strip()) < 300]
    
    if not sentences:
        return {"error": "Could not extract enough sentences to generate questions."}

    # Shuffle sentences to get random questions
    random.shuffle(sentences)
    
    # --- GENERATE MCQS (Cloze Deletion) ---
    mcqs = []
    num_mcqs = min(config['num_mcqs'], len(sentences))
    
    for i in range(num_mcqs):
        sentence = sentences[i]
        words = sentence.split()
        
        if len(words) < 4: continue
        
        # Try to find a "keyword" (longer than 4 chars, not a common stop word)
        # Simple heuristic: Capitalized words (nouns) or long words
        candidates = [w for w in words if len(w) > 4 and w.isalpha()]
        
        if not candidates:
            # Fallback to any long word
            candidates = [w for w in words if len(w) > 4]
            
        if not candidates: continue
            
        target_word = random.choice(candidates)
        target_index = words.index(target_word)
        
        # Create Question
        masked_words = words.copy()
        masked_words[target_index] = "______"
        question_text = " ".join(masked_words) + " (Fill in the blank)"
        
        # Create Distractors
        # Pick random words from other sentences
        distractors = set()
        while len(distractors) < 3:
            rand_sent = random.choice(sentences)
            rand_words = [w for w in rand_sent.split() if len(w) > 4]
            if rand_words:
                d = random.choice(rand_words)
                if d.lower() != target_word.lower():
                    distractors.add(d)
        
        options = list(distractors) + [target_word]
        random.shuffle(options) # Shuffle options
        
        # Map option to letter (A, B, C, D)
        option_map = {chr(65+i): opt for i, opt in enumerate(options)}
        correct_letter = [k for k, v in option_map.items() if v == target_word][0]
        
        mcqs.append({
            "question": question_text,
            "options": [f"{k}. {v}" for k, v in option_map.items()], # Format as "A. Option"
            "correct": correct_letter
        })

    # --- GENERATE SHORT QUESTIONS ---
    shorts = []
    num_short = min(config['num_short'], len(sentences) - num_mcqs)
    for i in range(num_mcqs, num_mcqs + num_short):
        sentence = sentences[i]
        # Strip trailing punctuation for the question
        clean_q = sentence.rstrip('.')
        shorts.append(f"Define or Explain: {clean_q}")

    # --- GENERATE LONG QUESTIONS ---
    longs = []
    num_long = min(config['num_long'], len(sentences) - (num_mcqs + num_short))
    for i in range(num_mcqs + num_short, num_mcqs + num_short + num_long):
        sentence = sentences[i]
        clean_q = sentence.rstrip('.')
        longs.append(f"Write a detailed note on: {clean_q}")

    # Construct JSON
    return {
        "subject": "Generated Exam",
        "class_level": config['academic_level'],
        "time": "2 Hours",
        "total_marks": str((num_mcqs * 1) + (num_short * 5) + (num_long * 10)),
        "section_a_mcqs": mcqs,
        "section_b_short": shorts,
        "section_c_long": longs
    }

# ==========================================
# EXPORT FUNCTIONS (UNCHANGED)
# ==========================================

def create_word_document(exam: Dict, show_answers: bool = False, is_answer_key_only: bool = False) -> BytesIO:
    doc = Document()
    title = "MCQ Answer Key" if is_answer_key_only else "Exam Paper"
    head = doc.add_heading(title, 0)
    head.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    p.add_run(f"Subject: {exam.get('subject', 'N/A')}\nClass: {exam.get('class_level', 'N/A')}\nTime: {exam.get('time', '2 Hours')}\n")
    
    if exam.get('section_a_mcqs'):
        doc.add_heading("SECTION A – MCQs", level=1)
        for i, q in enumerate(exam['section_a_mcqs'], 1):
            if is_answer_key_only:
                p = doc.add_paragraph(f"Q{i}) {q.get('correct', '')}")
            else:
                p = doc.add_paragraph(style='List Number')
                p.add_run(q.get('question', '')).bold = True
                for opt in q.get('options', []):
                    doc.add_paragraph(opt, style='List Bullet')
                if show_answers:
                     doc.add_paragraph(f"Answer: {q.get('correct', '')}").italic = True

    if not is_answer_key_only:
        if exam.get('section_b_short'):
            doc.add_heading("SECTION B – Short Questions", level=1)
            for i, q in enumerate(exam['section_b_short'], 1):
                p = doc.add_paragraph(style='List Number')
                p.add_run(q)
        if exam.get('section_c_long'):
            doc.add_heading("SECTION C – Long Questions", level=1)
            for i, q in enumerate(exam['section_c_long'], 1):
                p = doc.add_paragraph(style='List Number')
                p.add_run(q)
    
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def create_pdf_document(exam: Dict, show_answers: bool = False, is_answer_key_only: bool = False) -> BytesIO:
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=0.75*inch, rightMargin=0.75*inch)
    story = []
    styles = getSampleStyleSheet()
    style_normal = styles["BodyText"]
    style_normal.leading = 18
    style_header = ParagraphStyle('CustomHeader', parent=styles['Heading1'], alignment=1) 
    
    title_text = "MCQ Answer Key" if is_answer_key_only else "Examination Paper"
    story.append(Paragraph(title_text, style_header))
    info_text = f"<b>Subject:</b> {exam.get('subject', '')} | <b>Class:</b> {exam.get('class_level', '')}<br/>Time: {exam.get('time', '')} | Marks: {exam.get('total_marks', '')}"
    story.append(Paragraph(info_text, style_normal))
    story.append(Spacer(1, 12))

    if exam.get('section_a_mcqs'):
        story.append(Paragraph("SECTION A – MCQs", styles['Heading2']))
        for i, q in enumerate(exam['section_a_mcqs'], 1):
            if is_answer_key_only:
                text = f"{i}. {q.get('correct', '')}"
                story.append(Paragraph(text, style_normal))
            else:
                q_text = f"<b>Q{i}.</b> {q.get('question', '')}"
                story.append(Paragraph(q_text, style_normal))
                for opt in q.get('options', []):
                    story.append(Paragraph(f"• {opt}", style_normal))
                if show_answers:
                    story.append(Paragraph(f"<i>Correct Answer: {q.get('correct', '')}</i>", style_normal))
                story.append(Spacer(1, 6))

    if not is_answer_key_only:
        if exam.get('section_b_short'):
            story.append(Paragraph("SECTION B – Short Questions", styles['Heading2']))
            for i, q in enumerate(exam['section_b_short'], 1):
                story.append(Paragraph(f"<b>Q{i}.</b> {q}", style_normal))
                story.append(Spacer(1, 6))
        if exam.get('section_c_long'):
            story.append(Paragraph("SECTION C – Long Questions", styles['Heading2']))
            for i, q in enumerate(exam['section_c_long'], 1):
                story.append(Paragraph(f"<b>Q{i}.</b> {q}", style_normal))
                story.append(Spacer(1, 6))

    doc.build(story)
    buffer.seek(0)
    return buffer

# ==========================================
# FILE HELPERS (UNCHANGED)
# ==========================================

def extract_text_from_file(uploaded_file) -> str:
    text = ""
    try:
        if uploaded_file.type == "application/pdf":
            reader = pypdf.PdfReader(uploaded_file)
            for page in reader.pages:
                text += page.extract_text() + "\n"
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            doc = Document(uploaded_file)
            for para in doc.paragraphs:
                text += para.text + "\n"
        elif uploaded_file.type == "text/plain":
            text = str(uploaded_file.read(), "utf-8")
    except Exception as e:
        st.error(f"Error reading file: {e}")
    return re.sub(r'\s+', ' ', text).strip()

# ==========================================
# UI LOGIC
# ==========================================

def main():
    st.markdown('<div class="main-header">💻 Local AI Exam Generator</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-header">100% Offline | No API Keys Required</div>', unsafe_allow_html=True)

    with st.sidebar:
        st.header("⚙️ Configuration")
        st.info("Since this is the Local Version, we use a Rule-Based Engine to generate 'Fill in the Blank' questions directly from your text.")
        st.divider()

    # Step 1: Upload
    if st.session_state.step == 1:
        st.header("Step 1: Content Ingestion")
        uploaded_file = st.file_uploader("Upload Syllabus / Book", type=['pdf', 'docx', 'txt'])
        if uploaded_file:
            with st.spinner("Analyzing structure locally..."):
                st.session_state.full_text = extract_text_from_file(uploaded_file)
                if st.session_state.full_text:
                    # Use local heuristic analyzer
                    st.session_state.chapters_data = analyze_structure_locally(st.session_state.full_text)
                    st.session_state.step = 2
                    st.rerun()

    # Step 2: Select
    elif st.session_state.step == 2:
        st.header("Step 2: Select Chapters / Sections")
        st.info(f"Found {len(st.session_state.chapters_data)} sections.")
        
        # Extract just names for display
        options = [c['name'] for c in st.session_state.chapters_data]
        selected_indices = st.multiselect("Choose sections to test:", options=options, default=options)
        
        if st.button("Continue"):
            if selected_indices:
                # Store the actual content of selected chapters
                selected_content_list = [
                    c['content'] for c in st.session_state.chapters_data 
                    if c['name'] in selected_indices
                ]
                # Join content for the generation phase
                st.session_state.selected_text = "\n".join(selected_content_list)
                st.session_state.step = 3
                st.rerun()

    # Step 3: Config
    elif st.session_state.step == 3:
        st.header("Step 3: Exam Configuration")
        with st.form("config"):
            col1, col2 = st.columns(2)
            with col1:
                # Note: 'Difficulty' is visual only in local mode, as we just extract sentences
                exam_type = st.selectbox("Exam Type", ["Both", "MCQs only", "Subjective only"])
                diff = st.selectbox("Difficulty (Visual)", ["Mixed", "Easy", "Medium", "Hard"])
                lvl = st.selectbox("Level", ["School", "College", "University"])
            with col2:
                mcqs = st.slider("MCQs", 0, 20, 5)
                short = st.slider("Short", 0, 10, 3)
                long = st.slider("Long", 0, 5, 2)
            
            submitted = st.form_submit_button("Generate Locally")
            if submitted:
                with st.spinner("Generating questions from text..."):
                    cfg = {"exam_type": exam_type, "difficulty": diff, "academic_level": lvl, "num_mcqs": mcqs, "num_short": short, "num_long": long}
                    
                    # Pass the selected text to the local generator
                    result = local_generate_exam(st.session_state.selected_text, cfg)
                    
                    if result:
                        if "error" in result:
                            st.error(result["error"])
                        else:
                            st.session_state.generated_exam = result
                            st.session_state.step = 4
                            st.rerun()

    # Step 4: Download
    elif st.session_state.step == 4:
        exam = st.session_state.generated_exam
        st.header("Generated Exam Paper")
        
        # Display
        with st.container():
            st.markdown(f"""
            <div class="exam-paper">
                <div style="text-align: center; margin-bottom: 30px;">
                    <h2>{exam.get('subject', 'Subject')}</h2>
                    <p><strong>Class:</strong> {exam.get('class_level', '')} | <strong>Time:</strong> {exam.get('time', '')}</p>
                </div>
            """, unsafe_allow_html=True)
            
            if exam.get('section_a_mcqs'):
                st.markdown('<div class="section-header">SECTION A – MCQs</div>', unsafe_allow_html=True)
                for i, q in enumerate(exam['section_a_mcqs'], 1):
                    st.write(f"**Q{i}.** {q.get('question', '')}")
                    cols = st.columns(4)
                    for idx, opt in enumerate(q.get('options', [])):
                        cols[idx].text(opt)
            
            if exam.get('section_b_short'):
                st.markdown('<div class="section-header">SECTION B – Short Questions</div>', unsafe_allow_html=True)
                for i, q in enumerate(exam['section_b_short'], 1):
                    st.write(f"**Q{i}.** {q}")
            
            if exam.get('section_c_long'):
                st.markdown('<div class="section-header">SECTION C – Long Questions</div>', unsafe_allow_html=True)
                for i, q in enumerate(exam['section_c_long'], 1):
                    st.write(f"**Q{i}.** {q}")
            st.markdown("</div>", unsafe_allow_html=True)

        st.divider()
        
        # Download Section
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 📄 Question Paper (Student)")
            doc_word = create_word_document(exam, show_answers=False, is_answer_key_only=False)
            doc_pdf = create_pdf_document(exam, show_answers=False, is_answer_key_only=False)
            st.download_button("Download Word", data=doc_word, file_name="Question_Paper.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.download_button("Download PDF", data=doc_pdf, file_name="Question_Paper.pdf", mime="application/pdf")

        with col2:
            st.markdown("### 🗝️ MCQ Answer Key (Teacher)")
            key_word = create_word_document(exam, show_answers=True, is_answer_key_only=True)
            key_pdf = create_pdf_document(exam, show_answers=True, is_answer_key_only=True)
            st.download_button("Download Word", data=key_word, file_name="MCQ_Answer_Key.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.download_button("Download PDF", data=key_pdf, file_name="MCQ_Answer_Key.pdf", mime="application/pdf")

        if st.button("Start Over"):
            for key in list(st.session_state.keys()): del st.session_state[key]
            st.rerun()

if __name__ == "__main__":
    main()
