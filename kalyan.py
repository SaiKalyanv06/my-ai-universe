# =========================================================
# SAI'S AI UNIVERSE - COMPLETE ULTRA AI PLATFORM
# =========================================================

# =========================================================
# INSTALL THESE FIRST
# =========================================================

# pip install streamlit groq python-dotenv PyPDF2 gtts
# pip install streamlit-mic-recorder python-pptx python-docx
# pip install reportlab pillow requests speechrecognition
# pip install beautifulsoup4 lxml pandas youtube-transcript-api
# pip install pytz extra-streamlit-components

# =========================================================
# IMPORTS
# =========================================================

import streamlit as st
from groq import Groq
from dotenv import load_dotenv
from PyPDF2 import PdfReader
from gtts import gTTS
from streamlit_mic_recorder import mic_recorder

from pptx import Presentation
from pptx.util import Inches

from docx import Document
from reportlab.pdfgen import canvas

from PIL import Image

from bs4 import BeautifulSoup

import speech_recognition as sr
import pandas as pd
from youtube_transcript_api import YouTubeTranscriptApi

import tempfile
import requests
import sqlite3
import hashlib
import os
import base64
import datetime
import urllib.parse
import pytz
import extra_streamlit_components as stx

# =========================================================
# PAGE CONFIG - UPDATED WITH LOGO & NEW NAME
# =========================================================

st.set_page_config(
    page_title="Sai's AI Universe",  # Updated Name
    page_icon="logo.png",           # Updated Logo File
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================
# CUSTOM CSS (PRO UI FROM PHOTOS)
# =========================================================

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');

.stApp {
    background-color: #0B0F19;
    color: #FFFFFF;
    font-family: 'Inter', sans-serif;
}
[data-testid="stSidebar"] {
    background-color: #0E121D !important;
    border-right: 1px solid #1F2433 !important;
}
h1, h2, h3, h4, h5, h6, p, label {
    color: #FFFFFF !important;
}
.tool-card {
    background: #161B2A;
    border: 1px solid #23293B;
    border-radius: 16px;
    padding: 25px;
    text-align: left;
    transition: all 0.3s ease;
    cursor: pointer;
    height: 100%;
}
.tool-card:hover {
    border-color: #6C63FF;
    transform: translateY(-5px);
    box-shadow: 0 10px 20px rgba(108, 99, 255, 0.1);
}
.login-left {
    background: #11141B;
    padding: 60px;
    border-radius: 20px;
    border: 1px solid #1F2433;
}
.stButton button {
    background: linear-gradient(135deg, #6C63FF 0%, #8B5CF6 100%) !important;
    border-radius: 10px !important;
    border: none !important;
    font-weight: 600 !important;
    color: white !important;
    padding: 10px 20px !important;
    width: 100%;
}
.stTextInput input, .stTextArea textarea, .stSelectbox > div > div {
    background-color: #161B2A !important;
    border: 1px solid #23293B !important;
    border-radius: 10px !important;
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# LOAD ENV
# =========================================================

load_dotenv()

api_key = os.getenv("GROQ_API_KEY")

if not api_key:
    st.error("Please Add GROQ_API_KEY In .env File")
    st.stop()

# =========================================================
# CONNECT GROQ
# =========================================================

client = Groq(api_key=api_key)

# =========================================================
# DATABASE
# =========================================================

conn = sqlite3.connect(
    "users.db",
    check_same_thread=False
)

cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS users (
    username TEXT,
    password TEXT
)
""")

try:
    cursor.execute("ALTER TABLE users ADD COLUMN email TEXT")
except:
    pass

cursor.execute("""
CREATE TABLE IF NOT EXISTS vault (
    username TEXT, 
    filename TEXT, 
    filedata BLOB
)
""")

conn.commit()

# =========================================================
# PASSWORD HASH
# =========================================================

def hash_password(password):
    return hashlib.sha256(
        password.encode()
    ).hexdigest()

# =========================================================
# COOKIE MANAGER FOR PERSISTENT LOGIN 
# =========================================================

cookie_manager = stx.CookieManager(key="cookie_manager")
cookies_ready = cookie_manager.get_all()

# =========================================================
# WORLDWIDE LIVE SEARCH (PERFECT FIX)
# =========================================================

def get_live_search_data(query):
    context = ""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    # 1. SMART DATE CONVERTER (IST Fix)
    from datetime import datetime, timedelta
    
    # IST timezone define cheyyandi
    ist = pytz.timezone('Asia/Kolkata')
    
    # UTC time ni teesukuni, danni IST ki convert cheyyali
    now_utc = datetime.now(pytz.utc)
    now_ist = now_utc.astimezone(ist)
    
    today = now_ist.date()
    yesterday = today - timedelta(days=1)
    
    search_query = query.lower()
    if "yesterday" in search_query or "yesterdays" in search_query:
        search_query = search_query.replace("yesterdays", yesterday.strftime("%B %d, %Y")).replace("yesterday", yesterday.strftime("%B %d, %Y"))
    if "today" in search_query or "todays" in search_query:
        search_query = search_query.replace("todays", today.strftime("%B %d, %Y")).replace("today", today.strftime("%B %d, %Y"))
        
    # Layer 1: DuckDuckGo Lite
    try:
        url = "https://lite.duckduckgo.com/lite/"
        data = {"q": search_query}
        res = requests.post(url, headers=headers, data=data, timeout=6)
        if res.status_code == 200:
            soup = BeautifulSoup(res.text, 'html.parser')
            snippets = soup.find_all('td', class_='result-snippet')
            for snippet in snippets[:4]:
                context += f"- {snippet.text.strip()}\n"
            if context.strip():
                return context
    except:
        pass

    # Layer 2: Google Search Scrape
    try:
        google_url = f"https://www.google.com/search?q={urllib.parse.quote(search_query)}"
        res = requests.get(google_url, headers=headers, timeout=6)
        soup = BeautifulSoup(res.text, 'html.parser')
        for div in soup.find_all('div', class_=['BNeawe', 's3v9rd', 'AP7Wnd'])[:5]:
            text = div.get_text()
            if len(text) > 20:
                context += f"- {text.strip()}\n"
        if context.strip():
            return context
    except:
        pass

    # Layer 3: Wikipedia Fallback
    try:
        wiki_url = f"https://en.wikipedia.org/w/api.php?action=query&list=search&srsearch={urllib.parse.quote(search_query)}&utf8=&format=json"
        wiki_res = requests.get(wiki_url, timeout=5).json()
        if 'query' in wiki_res and 'search' in wiki_res['query']:
            for item in wiki_res['query']['search'][:3]:
                clean_text = BeautifulSoup(item['snippet'], 'html.parser').text
                context += f"- {item['title']}: {clean_text}\n"
    except:
        pass

    return context

# =========================================================
# STATE MANAGEMENT FOR MENU
# =========================================================

if "menu" not in st.session_state:
    st.session_state.menu = "🏠 Dashboard"

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# =========================================================
# AUTO-LOGIN PROCESS (PERSISTENT LOGIN)
# =========================================================

# 1. Check if cookie already exists
if not st.session_state.logged_in:
    if "saved_username" in cookies_ready:
        st.session_state.logged_in = True
        st.session_state.username = cookies_ready["saved_username"]

# =========================================================
# LOGIN SYSTEM (SIDE-BY-SIDE PRO UI)
# =========================================================

if not st.session_state.logged_in:

    st.markdown("<br><br>", unsafe_allow_html=True)
    l_col, r_col, _ = st.columns([1.2, 1, 0.2])
    
    with l_col:
        # UPDATED LOGIN SCREEN TITLE
        st.markdown("""
        <div class="login-left">
            <h1 style="font-size: 3.2rem; line-height: 1.2;">Welcome to <br><span style="color: #6C63FF;">Sai's AI Universe</span></h1>
            <p style="color: #A0A4B8; font-size: 1.2rem; margin-top: 15px;">Intelligent, secure, and built to help you achieve more every single day.</p>
        </div>
        """, unsafe_allow_html=True)

    with r_col:
        st.markdown("<h2>Welcome back</h2><p style='color: #A0A4B8;'>Sign in to your account to continue</p>", unsafe_allow_html=True)

        auth_option = st.radio(
            "Choose Option",
            [
                "Login",
                "Signup"
            ],
            horizontal=True
        )

        username = st.text_input("Username")

        email = st.text_input("Email")

        password = st.text_input(
            "Password",
            type="password"
        )

        # =====================================================
        # SIGNUP
        # =====================================================

        if auth_option == "Signup":

            if st.button("Create Account"):

                hashed = hash_password(password)

                cursor.execute(
                    "INSERT INTO users (username, email, password) VALUES (?, ?, ?)",
                    (username, email, hashed)
                )

                conn.commit()

                st.success("Account Created Successfully")

        # =====================================================
        # LOGIN
        # =====================================================

        else:

            if st.button("Login"):

                hashed = hash_password(password)

                cursor.execute(
                    "SELECT * FROM users WHERE username=? AND password=?",
                    (username, hashed)
                )

                result = cursor.fetchone()

                if result:

                    import datetime as dt
                    cookie_manager.set("saved_username", username, expires_at=dt.datetime.now() + dt.timedelta(days=30))

                    st.session_state.logged_in = True

                    st.session_state.username = username

                    st.success("Login Successful!")
                    
                    import time
                    time.sleep(1)
                    st.rerun()

                else:

                    st.error("Invalid Username Or Password")

# =========================================================
# MAIN APP
# =========================================================

else:

    # =====================================================
    # SIDEBAR - UPDATED WITH LOGO & TITLE
    # =====================================================

    try:
        st.sidebar.image("logo.png", use_container_width=True) # Added Logo in Sidebar
    except:
        pass
    
    st.sidebar.title("⚡ Sai's AI Universe")                # Updated Title

    st.sidebar.write(
        f"Welcome {st.session_state.username}"
    )

    st.sidebar.markdown("---")
    
    if st.sidebar.button("🏠 Home Dashboard", use_container_width=True):
        st.session_state.menu = "🏠 Dashboard"
        st.rerun()
        
    st.sidebar.markdown("---")


    model = st.sidebar.selectbox(
        "Choose Model",
        [
            "llama-3.3-70b-versatile",
            "mixtral-8x7b-32768",
            "gemma2-9b-it"
        ]
    )

    temperature = st.sidebar.slider(
        "Creativity",
        0.0,
        1.0,
        0.7
    )

    # =====================================================
    # MEMORY
    # =====================================================

    if "messages" not in st.session_state:

        st.session_state.messages = []

    # =====================================================
    # AI FUNCTION
    # =====================================================

    def generate_response(prompt):
        try:

            response = client.chat.completions.create(
                model=model,
                messages=[
                    {
                        "role": "user",
                        "content": prompt
                    }
                ],
                temperature=temperature
            )

            return response.choices[0].message.content
        except Exception as e:
            return f"Error: {e}"

    # =====================================================
    # DASHBOARD GRID UI (FROM PHOTOS)
    # =====================================================

    menu = st.session_state.menu

    if menu == "🏠 Dashboard":
        
        # INDIAN TIMEZONE FIX FOR GREETING
        ist = pytz.timezone('Asia/Kolkata')
        current_hour = datetime.datetime.now(ist).hour
        
        if current_hour < 12:
            greeting = "Good morning"
        elif current_hour < 18:
            greeting = "Good afternoon"
        else:
            greeting = "Good evening"

        st.markdown(f"<h1>{greeting}, {st.session_state.username} 👋</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #A0A4B8;'>What would you like to create or explore today?</p>", unsafe_allow_html=True)
        st.write("")

        tools = [
            ("💬 AI Chat", "Chat with intelligent AI"),
            ("💻 Coding Assistant", "Write and debug code"),
            ("📄 PDF Analyzer", "Extract info from PDFs"),
            ("🎨 AI Image Generator", "Create stunning images"),
            ("🔊 Voice Generator", "Text to speech conversion"),
            ("🧠 Study Assistant", "Learn any topic easily"),
            ("🌐 Translator", "Translate between languages"),
            ("📝 Resume Builder", "Create professional resumes"),
            ("📚 Quiz Generator", "Generate MCQ quizzes"),
            ("📧 Email Writer", "Draft professional emails"),
            ("📰 Article Writer", "Write detailed articles"),
            ("📱 Caption Generator", "Social media captions"),
            ("🧮 Math Solver", "Step-by-step math solutions"),
            ("📊 PPT Generator", "Create AI presentations"),
            ("📄 PDF Generator", "Create document PDFs"),
            ("📝 DOCX Generator", "Create Word documents"),
            ("🖼 Image To PDF", "Convert photos to PDF"),
            ("🤖 AI Agent", "Execute complex tasks"),
            ("📁 My Vault", "Save and manage your files"),
            ("👁️ Vision AI", "Analyze images with AI"),
            ("🌐 Web Summarizer", "Summarize webpage links"),
            ("🎙️ Audio To Text", "Convert audio to text"),
            ("📊 Data Analyzer", "Analyze CSV datasets"),
            ("🎬 YT Summarizer", "Summarize YouTube videos"),
            ("🛠️ Code Reviewer", "Fix and review code"),
            ("💡 Idea Validator", "Validate business ideas")
        ]

        for i in range(0, len(tools), 3):
            cols = st.columns(3)
            for j in range(3):
                if i + j < len(tools):
                    name, desc = tools[i+j]
                    with cols[j]:
                        st.markdown(f"""
                        <div class="tool-card">
                            <h3 style="margin-bottom:5px;">{name}</h3>
                            <p style="color:#A0A4B8; font-size:0.9rem;">{desc}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        if st.button(f"Open {name}", key=name, use_container_width=True):
                            st.session_state.menu = name
                            st.rerun()

    else:

        col1, col2 = st.columns([1, 5])
        with col1:
            if st.button("⬅ Back to Home"):
                st.session_state.menu = "🏠 Dashboard"
                st.rerun()
        
        st.markdown("---")

        if menu == "💬 AI Chat":

            st.title("🤖 AI Chat Assistant")

            for msg in st.session_state.messages:

                with st.chat_message(msg["role"]):

                    if isinstance(msg["content"], list):
                        for item in msg["content"]:
                            if item["type"] == "text":
                                st.markdown(item["text"])
                            elif item["type"] == "image_url":
                                st.image(item["image_url"]["url"])
                    else:
                        st.markdown(msg["content"])

            audio = mic_recorder(
                start_prompt="🎤 Start Recording",
                stop_prompt="⏹ Stop Recording",
                key="mic"
            )

            # GEMINI STYLE UI REDESIGN (Right aligned plus button)
            action_col1, action_col2 = st.columns([15, 1])

            with action_col2:
                with st.popover("➕"):  
                    st.markdown("**🛠️ Tools & Attachments**")
                    use_google = st.toggle("🌐 Web Search")
                    st.divider()
                    
                    attach_opt = st.radio("Attachment:", ["❌ None", "📁 Upload File", "📸 Open Camera"])
                    
                    attached_text = ""
                    attached_image_b64 = None
                    
                    if attach_opt == "📁 Upload File":
                        doc_file = st.file_uploader("Upload File", type=["png", "jpg", "jpeg", "txt", "pdf", "csv"])
                        if doc_file:
                            if doc_file.name.endswith((".png", ".jpg", ".jpeg")):
                                st.image(doc_file, width=200)
                                doc_file.seek(0)
                                attached_image_b64 = base64.b64encode(doc_file.read()).decode('utf-8')
                                st.success("Attached!")
                            elif doc_file.name.endswith(".txt"):
                                attached_text = f"\n\n[Attached File Content]:\n" + doc_file.read().decode("utf-8")
                                st.success("Attached!")
                            elif doc_file.name.endswith(".pdf"):
                                try:
                                    pdf_reader = PdfReader(doc_file)
                                    pdf_text = "".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
                                    attached_text = f"\n\n[Attached PDF Content]:\n" + pdf_text[:4000]
                                    st.success("Attached!")
                                except Exception as e:
                                    st.error(f"PDF Error: {e}")
                            elif doc_file.name.endswith(".csv"):
                                try:
                                    df_chat = pd.read_csv(doc_file)
                                    attached_text = f"\n\n[Attached CSV Data Sample]:\n" + df_chat.head(5).to_string()
                                    st.success("Attached!")
                                except Exception as e:
                                    st.error(f"CSV Error: {e}")
                                    
                    elif attach_opt == "📸 Open Camera":
                        st.info("💡 Camera is ON. Select 'None' to save battery.")
                        cam_file = st.camera_input("Take a photo")
                        if cam_file:
                            st.image(cam_file, width=200)
                            cam_file.seek(0)
                            attached_image_b64 = base64.b64encode(cam_file.read()).decode('utf-8')
                            st.success("Photo attached!")

            # The chat input sits cleanly below the right-aligned button
            prompt = st.chat_input("Ask Anything...")

            if prompt:
                
                # --- NEW FEATURE: AI IMAGE EDITING VIA PROMPT ---
                is_edit_request = attached_image_b64 and any(word in prompt.lower() for word in ["edit", "change", "make it", "filter", "black and white", "blur", "bright", "dark"])
                
                if is_edit_request:
                    user_message = {
                        "role": "user",
                        "content": [
                            {"type": "text", "text": prompt},
                            {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{attached_image_b64}"}}
                        ]
                    }
                    st.session_state.messages.append(user_message)
                    
                    with st.chat_message("user"):
                        st.markdown(prompt)
                        st.caption("📸 Image Attached for Editing")
                        
                    with st.chat_message("assistant"):
                        st.markdown("🎨 Processing your image edit request...")
                        try:
                            import io
                            from PIL import Image, ImageFilter, ImageEnhance
                            
                            img_data = base64.b64decode(attached_image_b64)
                            img = Image.open(io.BytesIO(img_data))
                            
                            # Simple logic to apply edits based on the prompt
                            if "black and white" in prompt.lower() or "grayscale" in prompt.lower():
                                img = img.convert("L")
                            elif "blur" in prompt.lower():
                                img = img.filter(ImageFilter.BLUR)
                            elif "bright" in prompt.lower():
                                enhancer = ImageEnhance.Brightness(img)
                                img = enhancer.enhance(1.5)
                            elif "dark" in prompt.lower():
                                enhancer = ImageEnhance.Brightness(img)
                                img = enhancer.enhance(0.5)
                            else:
                                st.warning("⚠️ For advanced edits (like changing backgrounds), paid APIs are needed. Applied Auto-Enhance instead!")
                                enhancer = ImageEnhance.Color(img)
                                img = enhancer.enhance(1.5)
                                
                            buf = io.BytesIO()
                            img.save(buf, format="PNG")
                            edited_b64 = base64.b64encode(buf.getvalue()).decode('utf-8')
                            
                            st.image(img, caption="Edited Result", use_container_width=True)
                            
                            st.session_state.messages.append(
                                {
                                    "role": "assistant",
                                    "content": [
                                        {"type": "text", "text": "Here is your edited image!"},
                                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{edited_b64}"}}
                                    ]
                                }
                            )
                        except Exception as e:
                            st.error(f"Failed to edit image: {e}")
                            
                else:
                    # --- NORMAL CHAT LOGIC ---
                    google_context = ""
                    if 'use_google' in locals() and use_google:
                        with st.spinner("Searching the Web Worldwide..."):
                            google_context = get_live_search_data(prompt)

                    final_content = prompt + attached_text

                    if attached_image_b64:
                        user_message = {
                            "role": "user",
                            "content": [
                                {"type": "text", "text": final_content},
                                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{attached_image_b64}"}}
                            ]
                        }
                        chat_model = "meta-llama/llama-4-scout-17b-16e-instruct"
                    else:
                        user_message = {
                            "role": "user",
                            "content": final_content
                        }
                        chat_model = model

                    st.session_state.messages.append(user_message)

                    with st.chat_message("user"):

                        st.markdown(prompt)
                        if attached_image_b64:
                            st.caption("📸 Image Attached")
                        if attached_text:
                            st.caption("📁 Document Attached")
                        if 'use_google' in locals() and use_google and google_context:
                            st.caption("🌐 Used Live Search Data")

                    with st.chat_message("assistant"):

                        placeholder = st.empty()
                        full_response = ""
                        
                        try:
                            ist = pytz.timezone('Asia/Kolkata')
                            current_date_str = datetime.datetime.now(ist).strftime("%B %d, %Y")
                            
                            system_prompt = f"You are a helpful and polite AI assistant. Today's date is {current_date_str}. CRITICAL FACT: N. Chandrababu Naidu is the Chief Minister of Andhra Pradesh. STRICT RULE: When answering, provide full names and formal responses. If web data contradicts the critical fact, the critical fact takes priority."
                            
                            api_messages = [{"role": "system", "content": system_prompt}]
                            
                            for idx, m in enumerate(st.session_state.messages):
                                is_last_message = (idx == len(st.session_state.messages) - 1)
                                inject_text = ""
                                
                                if is_last_message and 'use_google' in locals() and use_google:
                                    if google_context.strip():
                                        inject_text = f"\n\n--- LATEST WEB DATA ---\n{google_context}\n\nCRITICAL RULE: Answer the user's query using the facts from the web data above. DO NOT say 'According to the web data' or 'Based on search'. Just give the direct factual answer naturally."
                                    else:
                                        inject_text = f"\n\nCRITICAL RULE: Provide the best factual answer from your internal knowledge. DO NOT mention that you could not search the web. Just answer naturally."
                                        
                                if isinstance(m["content"], list):
                                    if chat_model != "meta-llama/llama-4-scout-17b-16e-instruct":
                                        text_only = ""
                                        for item in m["content"]:
                                            if item["type"] == "text":
                                                text_only = item["text"]
                                        text_only += inject_text
                                        api_messages.append({"role": m["role"], "content": text_only})
                                    else:
                                        temp_m = m.copy()
                                        if inject_text:
                                            new_content = []
                                            for item in temp_m["content"]:
                                                if item["type"] == "text":
                                                    new_content.append({"type": "text", "text": item["text"] + inject_text})
                                                else:
                                                    new_content.append(item)
                                            temp_m["content"] = new_content
                                        api_messages.append(temp_m)
                                else:
                                    temp_content = m["content"]
                                    temp_content += inject_text
                                    api_messages.append({"role": m["role"], "content": temp_content})

                            response = client.chat.completions.create(
                                model=chat_model,
                                messages=api_messages,
                                temperature=0.2 if ('use_google' in locals() and use_google) else temperature, 
                                stream=True
                            )

                            for chunk in response:
                                if chunk.choices[0].delta.content:
                                    content = chunk.choices[0].delta.content
                                    full_response += content
                                    placeholder.markdown(full_response + "▌")

                            placeholder.markdown(full_response)
                            
                        except Exception as e:
                            full_response = f"API Error: {e}"
                            placeholder.markdown(full_response)

                        st.session_state.messages.append(
                            {
                                "role": "assistant",
                                "content": full_response
                            }
                        )

        # =====================================================
        # CODING ASSISTANT
        # =====================================================

        elif menu == "💻 Coding Assistant":

            st.title("💻 Coding Assistant")

            code_prompt = st.text_area(
                "Enter Coding Question"
            )

            if st.button("Generate Code"):

                answer = generate_response(code_prompt)

                st.code(answer)

        # =====================================================
        # PDF ANALYZER
        # =====================================================

        elif menu == "📄 PDF Analyzer":

            st.title("📄 PDF Analyzer")

            uploaded_file = st.file_uploader(
                "Upload PDF",
                type="pdf"
            )

            if uploaded_file:

                pdf_reader = PdfReader(uploaded_file)

                text = "".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])

                st.success("PDF Loaded Successfully")

                question = st.text_input(
                    "Ask Question From PDF"
                )

                if st.button("Analyze PDF"):

                    prompt = f"""
                    PDF Content:
                    {text[:5000]}

                    Question:
                    {question}
                    """

                    answer = generate_response(prompt)

                    st.markdown(answer)

        # =====================================================
        # IMAGE GENERATOR
        # =====================================================

        elif menu == "🎨 AI Image Generator":

            st.title("🎨 AI Image Generator")

            image_prompt = st.text_input(
                "Describe Image"
            )

            if st.button("Generate Image"):

                formatted_prompt = image_prompt.replace(
                    " ",
                    "%20"
                )

                image_url = f"https://image.pollinations.ai/prompt/{formatted_prompt}"

                response = requests.get(image_url)

                image_path = "generated_image.jpg"

                with open(image_path, "wb") as f:

                    f.write(response.content)

                st.image(
                    image_path,
                    caption=image_prompt,
                    use_container_width=True
                )

                with open(image_path, "rb") as file:

                    st.download_button(
                        label="⬇ Download Image",
                        data=file,
                        file_name="ai_image.jpg",
                        mime="image/jpg"
                    )

        # =====================================================
        # VOICE GENERATOR
        # =====================================================

        elif menu == "🔊 Voice Generator":

            st.title("🔊 Voice Generator")

            voice_text = st.text_area(
                "Enter Text"
            )

            if st.button("Generate Voice"):

                tts = gTTS(
                    text=voice_text,
                    lang='en'
                )

                audio_file = "temp_voice.mp3"

                tts.save(audio_file)

                st.audio(audio_file)

        # =====================================================
        # STUDY ASSISTANT
        # =====================================================

        elif menu == "🧠 Study Assistant":

            st.title("🧠 Study Assistant")

            topic = st.text_input(
                "Enter Topic"
            )

            if st.button("Explain Topic"):

                answer = generate_response(
                    f"Explain {topic} simply"
                )

                st.markdown(answer)

        # =====================================================
        # TRANSLATOR
        # =====================================================

        elif menu == "🌐 Translator":

            st.title("🌐 Translator")

            text = st.text_area("Enter Text")

            language = st.selectbox(
                "Translate To",
                [
                    "Telugu",
                    "Hindi",
                    "French",
                    "Spanish"
                ]
            )

            if st.button("Translate"):

                answer = generate_response(
                    f"Translate to {language}: {text}"
                )

                st.markdown(answer)

        # =====================================================
        # RESUME BUILDER
        # =====================================================

        elif menu == "📝 Resume Builder":

            st.title("📝 Resume Builder")

            details = st.text_area(
                "Enter Resume Details"
            )

            if st.button("Generate Resume"):

                answer = generate_response(
                    f"Create professional resume: {details}"
                )

                st.markdown(answer)

        # =====================================================
        # QUIZ GENERATOR
        # =====================================================

        elif menu == "📚 Quiz Generator":

            st.title("📚 Quiz Generator")

            topic = st.text_input("Quiz Topic")

            if st.button("Generate Quiz"):

                answer = generate_response(
                    f"Create MCQ quiz about {topic}"
                )

                st.markdown(answer)

        # =====================================================
        # EMAIL WRITER
        # =====================================================

        elif menu == "📧 Email Writer":

            st.title("📧 Email Writer")

            topic = st.text_input("Email Topic")

            if st.button("Generate Email"):

                answer = generate_response(
                    f"Write professional email about {topic}"
                )

                st.markdown(answer)

        # =====================================================
        # ARTICLE WRITER
        # =====================================================

        elif menu == "📰 Article Writer":

            st.title("📰 Article Writer")

            topic = st.text_input("Article Topic")

            if st.button("Generate Article"):

                answer = generate_response(
                    f"Write detailed article about {topic}"
                )

                st.markdown(answer)

        # =====================================================
        # CAPTION GENERATOR
        # =====================================================

        elif menu == "📱 Caption Generator":

            st.title("📱 Caption Generator")

            topic = st.text_input("Post Topic")

            if st.button("Generate Caption"):

                answer = generate_response(
                    f"Generate captions for {topic}"
                )

                st.markdown(answer)

        # =====================================================
        # MATH SOLVER
        # =====================================================

        elif menu == "🧮 Math Solver":

            st.title("🧮 Math Solver")

            problem = st.text_input(
                "Enter Math Problem"
            )

            if st.button("Solve"):

                answer = generate_response(
                    f"Solve step by step: {problem}"
                )

                st.markdown(answer)

                st.latex(
                    r"x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}"
                )

        # =====================================================
        # PPT GENERATOR
        # =====================================================

        elif menu == "📊 PPT Generator":

            st.title("📊 AI PPT Generator")

            ppt_topic = st.text_input(
                "Enter PPT Topic"
            )

            slides_count = st.slider(
                "Slides",
                3,
                15,
                5
            )

            if st.button("Generate PPT"):

                prompt = f"""
                Create content for {slides_count} slides about:
                {ppt_topic}
                """

                answer = generate_response(prompt)

                prs = Presentation()

                slides_data = answer.split("Slide")

                for slide_data in slides_data:

                    if slide_data.strip():

                        slide_layout = prs.slide_layouts[6]

                        slide = prs.slides.add_slide(
                            slide_layout
                        )

                        title_box = slide.shapes.add_textbox(
                            Inches(0.5),
                            Inches(0.3),
                            Inches(8),
                            Inches(1)
                        )

                        title_box.text = ppt_topic

                        content_box = slide.shapes.add_textbox(
                            Inches(0.7),
                            Inches(1.5),
                            Inches(5),
                            Inches(3)
                        )

                        content_box.text = slide_data

                        image_url = f"https://image.pollinations.ai/prompt/{ppt_topic}"

                        response = requests.get(image_url)

                        image_path = "ppt_image.jpg"

                        with open(image_path, "wb") as f:

                            f.write(response.content)

                        slide.shapes.add_picture(
                            image_path,
                            Inches(5.5),
                            Inches(1.5),
                            width=Inches(3)
                        )

                ppt_file = f"{ppt_topic}.pptx"

                prs.save(ppt_file)

                with open(ppt_file, "rb") as file:

                    st.download_button(
                        label="⬇ Download PPT",
                        data=file,
                        file_name=ppt_file,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

        # =====================================================
        # PDF GENERATOR
        # =====================================================

        elif menu == "📄 PDF Generator":

            st.title("📄 PDF Generator")

            topic = st.text_input(
                "Enter PDF Topic"
            )

            if st.button("Generate PDF"):

                answer = generate_response(
                    f"Write notes about {topic}"
                )

                pdf_file = f"{topic}.pdf"

                c = canvas.Canvas(pdf_file)

                text_object = c.beginText(40, 800)

                for line in answer.split("\n"):

                    text_object.textLine(line)

                c.drawText(text_object)

                c.save()

                with open(pdf_file, "rb") as file:

                    st.download_button(
                        label="⬇ Download PDF",
                        data=file,
                        file_name=pdf_file,
                        mime="application/pdf"
                    )

        # =====================================================
        # DOCX GENERATOR
        # =====================================================

        elif menu == "📝 DOCX Generator":

            st.title("📝 DOCX Generator")

            topic = st.text_input(
                "Enter DOCX Topic"
            )

            if st.button("Generate DOCX"):

                answer = generate_response(
                    f"Write document about {topic}"
                )

                doc = Document()

                doc.add_heading(topic, level=1)

                doc.add_paragraph(answer)

                docx_file = f"{topic}.docx"

                doc.save(docx_file)

                with open(docx_file, "rb") as file:

                    st.download_button(
                        label="⬇ Download DOCX",
                        data=file,
                        file_name=docx_file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

        # =====================================================
        # IMAGE TO PDF
        # =====================================================

        elif menu == "🖼 Image To PDF":

            st.title("🖼 Image To PDF")

            uploaded_image = st.file_uploader(
                "Upload Image",
                type=["png", "jpg", "jpeg"]
            )

            if uploaded_image:

                st.image(uploaded_image)

                if st.button("Convert To PDF"):

                    image = Image.open(uploaded_image)

                    pdf_file = "converted_image.pdf"

                    rgb_image = image.convert("RGB")

                    rgb_image.save(pdf_file)

                    with open(pdf_file, "rb") as file:

                        st.download_button(
                            label="⬇ Download PDF",
                            data=file,
                            file_name=pdf_file,
                            mime="application/pdf"
                        )

        # =====================================================
        # AI AGENT
        # =====================================================

        elif menu == "🤖 AI Agent":

            st.title("🤖 AI Agent")

            task = st.text_area(
                "Give Task To AI Agent"
            )

            if st.button("Run Agent"):

                answer = generate_response(
                    f"""
                    You are an advanced AI agent.

                    Complete this task:
                    {task}
                    """
                )

                st.markdown(answer)

        # =====================================================
        # NEW FEATURES
        # =====================================================

        elif menu == "📁 My Vault":

            st.title("📁 My Vault")

            st.write("Securely store and retrieve your generated files.")

            v_col1, v_col2 = st.columns(2)

            with v_col1:

                st.markdown("### Upload to Vault")

                v_upload = st.file_uploader(
                    "Upload any file to Vault"
                )

                if st.button("Secure Save"):

                    if v_upload:

                        cursor.execute(
                            "INSERT INTO vault (username, filename, filedata) VALUES (?, ?, ?)", 
                            (st.session_state.username, v_upload.name, v_upload.read())
                        )

                        conn.commit()

                        st.success(f"{v_upload.name} securely saved!")

            with v_col2:

                st.markdown("### Your Saved Files")

                cursor.execute(
                    "SELECT filename, filedata FROM vault WHERE username=?", 
                    (st.session_state.username,)
                )

                saved_files = cursor.fetchall()

                if saved_files:

                    for fname, fdata in saved_files:

                        st.download_button(
                            label=f"⬇ Download {fname}", 
                            data=fdata, 
                            file_name=fname
                        )

                else:

                    st.info("Vault is empty.")

        elif menu == "👁️ Vision AI":

            st.title("👁️ Vision AI")

            img_file = st.file_uploader(
                "Upload Image for Analysis", 
                type=["png", "jpg", "jpeg"]
            )

            if img_file:

                st.image(img_file, width=300)

                img_prompt = st.text_input("Ask about this image:")

                if st.button("Analyze Image"):

                    img_file.seek(0)

                    base64_image = base64.b64encode(img_file.read()).decode('utf-8')

                    try:

                        vision_resp = client.chat.completions.create(
                            model="meta-llama/llama-4-scout-17b-16e-instruct",
                            messages=[{
                                "role": "user",
                                "content": [
                                    {"type": "text", "text": img_prompt or "Describe this image in detail."},
                                    {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"}}
                                ]
                            }],
                            temperature=0.7
                        )

                        st.success(vision_resp.choices[0].message.content)

                    except Exception as e:

                        st.error(f"Vision API Error. {e}")

        # =====================================================
        # WEB SUMMARIZER
        # =====================================================

        elif menu == "🌐 Web Summarizer":

            st.title("🌐 Web Summarizer")

            url = st.text_input("Enter Website URL")

            if st.button("Summarize"):

                try:

                    res = requests.get(url)
                    
                    soup = BeautifulSoup(res.text, 'html.parser')
                    
                    text = soup.get_text(separator=' ', strip=True)[:5000]

                    prompt = f"Summarize this webpage content in detail:\n\n{text}"
                    
                    answer = generate_response(prompt)
                    
                    st.markdown(answer)

                except Exception as e:
                    
                    st.error(f"Error fetching URL: {e}")

        # =====================================================
        # AUDIO TO TEXT
        # =====================================================

        elif menu == "🎙️ Audio To Text":

            st.title("🎙️ Audio To Text")

            st.write("Upload a .wav audio file to extract text.")

            audio_file = st.file_uploader("Upload Audio", type=["wav"])

            if audio_file is not None:

                st.audio(audio_file)

                if st.button("Extract Text"):

                    recognizer = sr.Recognizer()
                    
                    with sr.AudioFile(audio_file) as source:

                        audio_data = recognizer.record(source)

                        try:

                            text = recognizer.recognize_google(audio_data)
                            
                            st.success("Extraction Successful!")
                            
                            st.write(text)

                        except Exception as e:
                            
                            st.error(f"Error converting audio: {e}")

        # =====================================================
        # DATA ANALYZER
        # =====================================================

        elif menu == "📊 Data Analyzer":

            st.title("📊 Data Analyzer")

            csv_file = st.file_uploader("Upload CSV File", type=["csv"])

            if csv_file is not None:

                df = pd.read_csv(csv_file)
                
                st.dataframe(df.head())

                question = st.text_input("Ask a question about this data")

                if st.button("Analyze Data"):

                    data_sample = df.head(10).to_string()
                    
                    data_info = df.describe().to_string()
                    
                    prompt = f"Here is a sample of my dataset:\n{data_sample}\n\nHere is the summary info:\n{data_info}\n\nBased on this, answer the user's question: {question}"

                    answer = generate_response(prompt)
                    
                    st.markdown(answer)

        # =====================================================
        # YT SUMMARIZER
        # =====================================================

        elif menu == "🎬 YT Summarizer":

            st.title("🎬 YouTube Video Summarizer")

            yt_url = st.text_input("Enter YouTube Video Link")

            if st.button("Summarize Video"):

                try:

                    if "v=" in yt_url:
                        
                        video_id = yt_url.split("v=")[1].split("&")[0]
                        
                    elif "youtu.be/" in yt_url:
                        
                        video_id = yt_url.split("youtu.be/")[1].split("?")[0]
                        
                    else:
                        
                        st.error("Invalid YouTube URL")
                        
                        video_id = None

                    if video_id:

                        try:
                            transcript_list = YouTubeTranscriptApi().fetch(video_id)
                        except AttributeError:
                            transcript_list = YouTubeTranscriptApi.get_transcript(video_id)
                        
                        transcript_text = " ".join([t.text if hasattr(t, 'text') else t['text'] for t in transcript_list])[:5000]

                        prompt = f"Summarize this YouTube video transcript in detail:\n\n{transcript_text}"
                        
                        answer = generate_response(prompt)
                        
                        st.markdown(answer)

                except Exception as e:
                    
                    st.error(f"Could not fetch transcript. The video might not have subtitles. Error: {e}")

        # =====================================================
        # CODE REVIEWER
        # =====================================================

        elif menu == "🛠️ Code Reviewer":

            st.title("🛠️ Code Bug Fixer & Reviewer")

            bug_code = st.text_area("Paste your code here with errors")

            if st.button("Review & Fix Code"):

                prompt = f"Act as an expert software engineer. Review the following code, find any bugs, explain them, and provide the corrected code:\n\n{bug_code}"

                answer = generate_response(prompt)
                
                st.markdown(answer)

        # =====================================================
        # IDEA VALIDATOR
        # =====================================================

        elif menu == "💡 Idea Validator":

            st.title("💡 Startup Idea Validator")

            st.write("Turn your raw business ideas into a complete execution plan.")

            idea = st.text_area("Describe your startup or business idea:")

            if st.button("Validate Idea"):

                if idea:

                    prompt = f"""
                    Act as an expert Silicon Valley Startup Consultant and Investor.
                    Analyze this business idea: "{idea}"
                    
                    Please provide a detailed report with the following sections:
                    1. 🎯 Elevator Pitch (1-2 lines)
                    2. 👥 Target Audience & Market Size
                    3. ⚔️ Competitor Analysis (Direct & Indirect)
                    4. 💰 Monetization Strategy (How to make money)
                    5. 🚦 SWOT Analysis (Strengths, Weaknesses, Opportunities, Threats)
                    6. 🚀 30-Day Action Plan (Step-by-step roadmap to launch)
                    """

                    answer = generate_response(prompt)
                    
                    st.markdown(answer)

                else:

                    st.warning("Please enter your idea first.")

    # =====================================================
    # LOGOUT
    # =====================================================
    
    st.sidebar.markdown("---")
    
    if st.sidebar.button("Logout"):

        st.session_state.logged_in = False
        cookie_manager.delete("saved_username") # Deleting Cookie on Logout

        st.rerun()
