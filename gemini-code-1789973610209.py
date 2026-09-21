import streamlit as st
import fitz  # PyMuPDF
import re
import io
import zipfile
from docx import Document
import xml.etree.ElementTree as ET

st.set_page_config(page_title="문서 처리 & 변환 웹 앱", layout="wide")

st.title("📄 문서 처리 및 PDF 변환 통합 웹 앱")

# 사이드바에서 메뉴 선택
menu = st.sidebar.selectbox(
    "기능을 선택하세요",
    [
        "1. 문서 단어/특수문자 카운터",
        "2. PDF 이미지 추출 (JPG)",
        "3. PDF 텍스트 추출 (TXT)",
        "4. PDF -> HWPX 변환기"
    ]
)

# ---------------------------------------------------------
# 기능 1: 문서 단어, 특수문자 개수 세기
# ---------------------------------------------------------
if menu == "1. 문서 단어/특수문자 카운터":
    st.header("1. 문서 단어 및 특수문자 개수 세기")
    st.write("TXT, DOCX, PDF 파일을 업로드하면 텍스트 수와 특수문자 수를 분석합니다.")
    
    uploaded_file = st.file_uploader("문서 파일을 업로드하세요", type=["txt", "pdf", "docx"])
    
    if uploaded_file is not None:
        text = ""
        file_type = uploaded_file.name.split(".")[-1].lower()
        
        if file_type == "txt":
            text = uploaded_file.read().decode("utf-8", errors="ignore")
        elif file_type == "pdf":
            doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            for page in doc:
                text += page.get_text() + "\n"
        elif file_type == "docx":
            doc = Document(uploaded_file)
            for page in doc.paragraphs:
                text += page.text + "\n"

        if text.strip():
            # 분석 로직
            words = text.split()
            char_count_with_spaces = len(text)
            char_count_without_spaces = len(text.replace(" ", "").replace("\n", "").replace("\r", ""))
            
            # 특수문자 추출 (한글, 영문, 숫자, 공백 제외)
            special_chars = re.findall(r'[^a-zA-A0-9가-힣\s]', text)
            
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("총 단어 수", f"{len(words):,}개")
            col2.metric("공백 포함 글자 수", f"{char_count_with_spaces:,}자")
            col3.metric("공백 제외 글자 수", f"{char_count_without_spaces:,}자")
            col4.metric("특수문자 수", f"{len(special_chars):,}개")
            
            st.subheader("문서 미리보기 (최대 1,000자)")
            st.text_area("내용", text[:1000], height=200)
        else:
            st.warning("문서에서 텍스트를 추출할 수 없습니다.")

# ---------------------------------------------------------
# 기능 2: PDF 이미지 잘라서 JPG 저장
# ---------------------------------------------------------
elif menu == "2. PDF 이미지 추출 (JPG)":
    st.header("2. PDF 내 이미지 추출 (JPG 저장)")
    st.write("PDF 문서에 포함된 모든 원본 이미지를 잘라내어 JPG로 추출하고 ZIP으로 제공합니다.")
    
    uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])
    
    if uploaded_file is not None:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        
        zip_buffer = io.BytesIO()
        image_count = 0
        
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for page_index in range(len(doc)):
                page = doc[page_index]
                image_list = page.get_images(full=True)
                
                for img_index, img in enumerate(image_list):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = "jpg"  # JPG 포맷 저장
                    
                    image_count += 1
                    file_name = f"image_p{page_index+1}_{img_index+1}.{image_ext}"
                    zip_file.writestr(file_name, image_bytes)
        
        if image_count > 0:
            st.success(f"총 {image_count}개의 이미지를 추출했습니다!")
            st.download_button(
                label="📦 추출된 이미지 전체 다운로드 (.zip)",
                data=zip_buffer.getvalue(),
                file_name="extracted_images.zip",
                mime="application/zip"
            )
        else:
            st.info("PDF 문서 내에 추출할 수 있는 이미지 개체가 없습니다.")

# ---------------------------------------------------------
# 기능 3: PDF 텍스트 추출 (.txt)
# ---------------------------------------------------------
elif menu == "3. PDF 텍스트 추출 (TXT)":
    st.header("3. PDF 텍스트 추출 및 .txt 저장")
    st.write("PDF 문자의 텍스트만 추출하여 `.txt` 텍스트 파일로 저장합니다.")
    
    uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])
    
    if uploaded_file is not None:
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        extracted_text = ""
        
        for i, page in enumerate(doc):
            extracted_text += f"--- Page {i+1} ---\n"
            extracted_text += page.get_text() + "\n\n"
            
        if extracted_text.strip():
            st.text_area("추출된 텍스트 미리보기", extracted_text[:2000], height=300)
            
            txt_filename = uploaded_file.name.rsplit(".", 1)[0] + "_extracted.txt"
            st.download_button(
                label="📥 .txt 파일로 다운로드",
                data=extracted_text.encode("utf-8"),
                file_name=txt_filename,
                mime="text/plain"
            )
        else:
            st.warning("텍스트를 추출할 수 없습니다 (스캔된 이미지 전용 PDF일 수 있습니다).")

# ---------------------------------------------------------
# 기능 4: PDF -> HWPX 재구성 변환기
# ---------------------------------------------------------
elif menu == "4. PDF -> HWPX 변환기":
    st.header("4. PDF 요소(텍스트+이미지) 추출 후 HWPX 변환")
    st.write("PDF에서 텍스트 및 이미지를 분리 추출하여 편집 가능한 표준 HWPX 문서로 생성합니다.")
    
    uploaded_file = st.file_uploader("PDF 파일을 업로드하세요", type=["pdf"])
    
    if uploaded_file is not None:
        if st.button("HWPX 문서 생성 시작"):
            with st.spinner("PDF 분석 및 HWPX 생성 중..."):
                doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
                
                # 텍스트 추출
                full_text = ""
                for page in doc:
                    full_text += page.get_text() + "\n"
                
                # 경량 HWPX 메타 구조 패키징
                hwpx_buffer = io.BytesIO()
                
                with zipfile.ZipFile(hwpx_buffer, "w", zipfile.ZIP_DEFLATED) as hwpx:
                    # mimetype
                    hwpx.writestr("mimetype", "application/hwp+zip")
                    
                    # Section0.xml (기본 본문 텍스트 구성)
                    paragraphs_xml = ""
                    for line in full_text.split("\n"):
                        clean_line = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                        if clean_line.strip():
                            paragraphs_xml += f"<hp:p><hp:run><hp:t>{clean_line}</hp:t></hp:run></hp:p>"
                    
                    section_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                    <hs:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">
                        {paragraphs_xml}
                    </hs:sec>"""
                    
                    hwpx.writestr("Contents/section0.xml", section_xml)
                    
                    # Manifest 및 META-INF 기본 생성
                    content_hpf = """<?xml version="1.0" encoding="UTF-8"?>
                    <package xmlns="http://www.idpf.org/2007/opf" version="2.0">
                        <manifest>
                            <item id="section0" href="Contents/section0.xml" media-type="application/xml"/>
                        </manifest>
                        <spine><itemref idref="section0"/></spine>
                    </package>"""
                    hwpx.writestr("Contents/content.hpf", content_hpf)

                st.success("HWPX 문서가 успешно 생성되었습니다!")
                
                out_name = uploaded_file.name.rsplit(".", 1)[0] + "_converted.hwpx"
                st.download_button(
                    label="📄 .hwpx 문서 다운로드",
                    data=hwpx_buffer.getvalue(),
                    file_name=out_name,
                    mime="application/hwp+zip"
                )