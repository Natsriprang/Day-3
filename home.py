import os
from io import BytesIO
from datetime import datetime

import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from docx import Document
from docx.shared import Pt

load_dotenv()

# 1) ตั้งค่า OpenAI client
#    รองรับทั้งการรัน local (env var) และบน Streamlit Cloud (st.secrets)
def get_api_key():
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.environ.get("OPENAI_API_KEY")


api_key = get_api_key()
client = OpenAI(api_key=api_key) if api_key else None

ALLOWED_MODELS = ["gpt-5-pro", "gpt-4o", "gpt-4o-mini"]
DEFAULT_MODEL = "gpt-4o"

# 2) หลักองค์ประกอบของสัญญา (Elements of a Contract) — สรุปจาก The Law Handbook
CONTRACT_ELEMENTS_KNOWLEDGE = """
คุณคือผู้ช่วยกฎหมาย (legal assistant) ที่เชี่ยวชาญเรื่อง "การเกิดขึ้นของสัญญา" (Contract Formation)
ภายใต้กฎหมายทั่วไปของออสเตรเลีย (โดยเฉพาะรัฐวิกตอเรีย) อ้างอิงจาก The Law Handbook
บทที่ 7.1 "Elements of a Contract" ซึ่งมีองค์ประกอบหลัก 6 ข้อดังนี้:

1) Offer and Acceptance (คำเสนอและคำสนอง)
   - สัญญาเกิดขึ้นเมื่อคำเสนอ (offer) ของฝ่ายหนึ่งได้รับการสนอง (accepted) จากอีกฝ่ายหนึ่ง
   - คำเสนอต้องเป็นคำมั่นที่ชัดเจนและแน่นอนว่าจะผูกพัน ไม่ใช่เพียงการแสดงความเต็มใจจะเจรจาต่อ
   - การสนองต้องตรงกับสิ่งที่เสนอเป๊ะๆ (mirror image) หากเปลี่ยนแปลงเงื่อนไข จะถือเป็น "คำเสนอโต้กลับ"
     (counter-offer) ซึ่งทำให้คำเสนอเดิมสิ้นผล และฝ่ายที่เสนอเดิมต้องเลือกรับหรือปฏิเสธคำเสนอโต้กลับนั้น
   - ผู้เสนอสามารถถอนคำเสนอได้ก่อนมีการสนองรับ แต่ต้องสื่อสารการถอนนั้นไปยังอีกฝ่ายอย่างชัดเจน
   - การสนองรับต้องชัดแจ้งและสื่อสารไปถึงผู้เสนอ กฎหมายไม่ถือว่าการนิ่งเฉยคือการยอมรับ
     แต่การสนองรับอาจเกิดจากพฤติกรรม (conduct) ได้เช่นกัน

2) Intention to Create Legal Relations (เจตนาที่จะผูกพันทางกฎหมาย)
   - แค่มีข้อตกลงกันยังไม่พอ คู่สัญญาต้องมีเจตนาที่จะให้ข้อตกลงนั้นมีผลผูกพันทางกฎหมาย
   - พิจารณาจากสิ่งที่ "บุคคลที่มีเหตุผล" (reasonable person) จะเข้าใจจากสถานการณ์ ไม่ใช่เจตนาในใจจริงๆ ของคู่กรณี
   - ในการทำธุรกิจ/การค้าทั่วไป (commercial/arm's length) กฎหมายจะสันนิษฐานไว้ก่อนว่ามีเจตนาผูกพันทางกฎหมาย
   - ในความสัมพันธ์ทางครอบครัวหรือเพื่อนฝูง/ข้อตกลงเชิงสังคม มักสันนิษฐานว่า "ไม่มี" เจตนาผูกพันทางกฎหมาย
     เว้นแต่จะแสดงเจตนาชัดเจนเป็นอย่างอื่น (เช่น ทำเป็นลายลักษณ์อักษร)

3) Consideration (ค่าตอบแทน/สิ่งตอบแทน)
   - คือ "ราคา" ที่จ่ายเพื่อแลกกับคำมั่นของอีกฝ่าย ไม่จำเป็นต้องเป็นเงิน แต่ต้องมีมูลค่าบางอย่าง
   - ศาลจะไม่ตัดสินความ "เหมาะสม" ของมูลค่า ตราบใดที่มีมูลค่าจริงอยู่ แม้จะเป็นจำนวนเพียงเล็กน้อยก็ตาม
   - ความรัก ความเสน่หา หรือของกำนัลโดยสมัครใจ (gift) ไม่ถือเป็น consideration ที่สมบูรณ์
   - ข้อยกเว้น: เอกสารที่ทำในรูปแบบ "deed" (นิติกรรมประทับตรา) ไม่จำเป็นต้องมี consideration

4) Legal Capacity (ความสามารถตามกฎหมายในการทำสัญญา)
   - บุคคลบางกลุ่มอาจมีข้อจำกัดในการทำสัญญาผูกพัน เช่น
     ผู้มีความบกพร่องทางจิตหรือถูกครอบงำจากยา/แอลกอฮอล์ชั่วคราว, ผู้เยาว์ (อายุต่ำกว่า 18 ปี),
     บุคคลล้มละลาย, นิติบุคคล/บริษัท (ต้องกระทำผ่านผู้มีอำนาจ), และผู้ต้องขัง
   - ผู้เยาว์: สัญญาเกี่ยวกับ "ของจำเป็น" (necessaries) มักผูกพันได้ ส่วนสัญญาสำหรับสิ่งที่ไม่ใช่ของจำเป็น
     หรือสัญญากู้ยืมเงิน มักถือเป็นโมฆะ (void) ไม่ผูกพันผู้เยาว์
   - ผู้มีความบกพร่องทางจิต: สัญญาจะไม่ผูกพันหากพิสูจน์ได้ว่าไม่เข้าใจลักษณะทั่วไปของสัญญา
     และอีกฝ่ายรู้หรือควรรู้ถึงความบกพร่องนั้น

5) Consent (ความยินยอมโดยแท้จริง)
   - การทำสัญญาต้องเกิดจากเจตจำนงเสรีและความเข้าใจที่แท้จริงของทั้งสองฝ่าย
   - ความยินยอมอาจถูกกระทบโดย: ความสำคัญผิด (mistake), การหลอกลวง/แสดงข้อความเท็จ (misrepresentation),
     การข่มขู่บังคับ (duress), การใช้อิทธิพลครอบงำโดยไม่เป็นธรรม (undue influence/unconscionability),
     หรือข้อสัญญาที่ไม่เป็นธรรมในสัญญาสำเร็จรูป (unfair terms in standard form contracts)
   - หากความยินยอมถูกกระทบอย่างมีนัยสำคัญ สัญญาอาจเป็นโมฆียะ (voidable) คือฝ่ายที่ถูกกระทบเลือกยกเลิกได้

6) Legality (ความชอบด้วยกฎหมายของสัญญา)
   - สัญญาที่มีวัตถุประสงค์ผิดกฎหมาย หรือขัดต่อนโยบายสาธารณะ (เช่น เพื่อกระทำความผิดอาญา,
     ขัดศีลธรรมทางเพศ, ขัดต่อการบริหารงานยุติธรรม, จำกัดการประกอบอาชีพการค้าโดยไม่สมเหตุผล ฯลฯ)
     จะถือเป็นโมฆะและไม่สามารถบังคับได้ตามกฎหมาย

หลักการวิเคราะห์ที่ต้องทำ
--------------------------
เมื่อผู้ใช้ป้อนเหตุการณ์เข้ามา ให้คุณ:
1. วิเคราะห์เหตุการณ์นั้นเทียบกับองค์ประกอบทั้ง 6 ข้อข้างต้น "ทีละข้อ" (ข้อที่ไม่เกี่ยวข้องกับข้อเท็จจริง
   ที่ให้มา ให้ระบุสั้นๆ ว่า "ไม่มีข้อเท็จจริงเพียงพอที่จะวิเคราะห์" หรือ "ดูเหมือนจะไม่มีปัญหาในประเด็นนี้")
2. สำหรับแต่ละองค์ประกอบที่เกี่ยวข้อง ให้อธิบายหลักกฎหมายสั้นๆ ก่อน แล้วจึงปรับใช้ (apply) กับข้อเท็จจริง
   ในรูปแบบ: "หลักกฎหมายคือ ... ดังนั้น ในกรณีนี้ ... เพราะ ..."
3. สรุปในตอนท้ายอย่างชัดเจนว่า "สัญญาเกิดขึ้นแล้วหรือยังไม่เกิดขึ้น" (Contract formed / Not yet formed /
   Uncertain - need more facts) พร้อมเหตุผลสั้นๆ
4. หากข้อมูลไม่เพียงพอที่จะสรุปในบางประเด็น ให้บอกตรงๆ ว่าต้องการข้อเท็จจริงเพิ่มเติมอะไร
5. ตอบเป็นภาษาเดียวกับที่ผู้ใช้พิมพ์เข้ามา (ถ้าผู้ใช้พิมพ์ภาษาไทย ให้ตอบภาษาไทย ถ้าพิมพ์อังกฤษ ให้ตอบอังกฤษ)
6. คำตอบต้องอิงตามหลักกฎหมายของออสเตรเลีย (Victoria) ตามที่สรุปไว้ข้างต้นเท่านั้น ห้ามอ้างกฎหมายประเทศอื่น
7. ระบุด้วยว่านี่เป็นข้อมูลเพื่อการศึกษาเบื้องต้นเท่านั้น ไม่ใช่คำปรึกษาทางกฎหมายอย่างเป็นทางการ
   (This is general information, not formal legal advice) ไว้สั้นๆ ท้ายคำตอบ

หากผู้ใช้ถามคำถามต่อเนื่อง (follow-up question) หลังจากที่คุณวิเคราะห์เหตุการณ์แรกไปแล้ว
ให้ตอบโดยอ้างอิงกับเหตุการณ์และบทวิเคราะห์เดิมในบทสนทนานี้ด้วย อย่าตอบราวกับว่าเป็นคำถามใหม่ที่ไม่มีบริบท
"""

SYSTEM_PROMPT = CONTRACT_ELEMENTS_KNOWLEDGE


# ---------------------------------------------------------------------------
# 3) ฟังก์ชันเรียก OpenAI API
#    รับ "ประวัติการสนทนาทั้งหมด" (ไม่ใช่แค่ข้อความล่าสุด) เพื่อให้ follow-up
#    คำถามยังมีบริบทของเหตุการณ์แรกและคำตอบก่อนหน้าอยู่ด้วย
# ---------------------------------------------------------------------------
def call_openai(history: list, model_name: str) -> str:
    messages = [{"role": "system", "content": SYSTEM_PROMPT}] + history
    response = client.chat.completions.create(
        model=model_name,
        messages=messages,
        temperature=0.2,
    )
    return response.choices[0].message.content


def build_initial_prompt(scenario_text: str) -> str:
    return (
        "Please analyze the situation based on the 6 elements of contract formation"
        "and summarize:\n\n"
        f"situation: {scenario_text}"
    )


def build_docx(messages: list) -> BytesIO:
    """
    แปลงประวัติการสนทนา (list ของ {"role", "content"}) ให้เป็นไฟล์ .docx
    แล้วคืนค่าเป็น BytesIO เพื่อส่งให้ st.download_button ใช้งาน
    """
    doc = Document()

    title = doc.add_heading("Contract Formation Chatbot — Discussion Log", level=1)
    subtitle = doc.add_paragraph(
        f"Exported on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    )
    subtitle.runs[0].italic = True
    doc.add_paragraph(
        "Reference: The Law Handbook (Victoria, Australia) — Elements of a Contract"
    )
    doc.add_paragraph()  # เว้นบรรทัด

    role_labels = {"user": "You", "assistant": "Assistant"}

    for msg in messages:
        role = role_labels.get(msg["role"], msg["role"].capitalize())
        heading = doc.add_paragraph()
        run = heading.add_run(f"{role}:")
        run.bold = True
        run.font.size = Pt(12)

        # เนื้อหาอาจมีหลายบรรทัด ให้แยกเป็นย่อหน้าตาม newline เพื่อให้อ่านง่ายใน Word
        for line in msg["content"].split("\n"):
            doc.add_paragraph(line if line.strip() else "")

        doc.add_paragraph()  # เว้นบรรทัดระหว่างข้อความแต่ละคน

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# 4) หน้าเว็บ Streamlit
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Contract Formation Chatbot (Australia)", page_icon="⚖️")

st.title("⚖️ Contract Formation Chatbot (Australia)")
st.caption(
    "This chatbot will answer your question about the contract formation in "
    "Victoria, Australia only, using the reference from The Law Handbook — "
    "Elements of a Contract."
)

if not api_key:
    st.error(
        "ไม่พบ OPENAI_API_KEY — ถ้ารัน local ให้ตั้งค่า environment variable, "
        "ถ้า deploy บน Streamlit Cloud ให้ใส่ใน App settings > Secrets"
    )

with st.sidebar:
    st.subheader("Settings")
    model_name = st.selectbox("Chatbot Models", ALLOWED_MODELS, index=ALLOWED_MODELS.index(DEFAULT_MODEL))
    if st.button("Clear discussion history"):
        st.session_state.messages = []
        st.session_state.scenario_submitted = False

if "messages" not in st.session_state:
    st.session_state.messages = []  # list of {"role": ..., "content": ...}
if "scenario_submitted" not in st.session_state:
    # True หลังจากผู้ใช้ submit เหตุการณ์แรกแล้ว ใช้ควบคุมว่าจะโชว์ช่อง follow-up หรือไม่
    st.session_state.scenario_submitted = False

# แสดงประวัติแชทที่ผ่านมา
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# ---------------------------------------------------------------------------
# ฟอร์มแรก: เหตุการณ์เริ่มต้น + ปุ่ม Submit
# ใช้ st.form เพื่อให้ AI ประมวลผล "ต่อเมื่อกด Submit เท่านั้น"
# แสดงเฉพาะตอนที่ยังไม่เคยส่งเหตุการณ์แรก (กันไม่ให้ผู้ใช้ยิงเหตุการณ์ใหม่ปนกับ follow-up)
# ---------------------------------------------------------------------------
if not st.session_state.scenario_submitted:
    with st.form(key="scenario_form", clear_on_submit=True):
        scenario_text = st.text_area(
            "Please provide your situation here.",
            placeholder="i.e. 'I offerred him 500$ for this camera but he did not agree.'",
            height=120,
        )
        submitted = st.form_submit_button("Submit")

    if submitted:
        if not scenario_text or not scenario_text.strip():
            st.warning("Please provide the situation before submit.")
        elif not api_key:
            st.warning("ยังไม่ได้ตั้งค่า OPENAI_API_KEY จึงเรียก AI ไม่ได้")
        else:
            user_prompt = build_initial_prompt(scenario_text)
            st.session_state.messages.append({"role": "user", "content": user_prompt})
            with st.chat_message("user"):
                st.markdown(scenario_text)

            with st.chat_message("assistant"):
                with st.spinner("assessing..."):
                    try:
                        answer = call_openai(st.session_state.messages, model_name)
                    except Exception as e:
                        answer = f"เกิดข้อผิดพลาดในการเรียก OpenAI API: {e}"
                    st.markdown(answer)

            st.session_state.messages.append({"role": "assistant", "content": answer})
            st.session_state.scenario_submitted = True
            st.rerun()

# ---------------------------------------------------------------------------
# ช่อง follow-up question — โชว์เฉพาะหลังจากมีการวิเคราะห์เหตุการณ์แรกแล้ว
# st.chat_input วางนอก st.form ไม่ได้ครอบด้วยฟอร์ม (ข้อจำกัดของ Streamlit)
# แต่ยังคงพฤติกรรม "ส่งเมื่อกด Enter/ปุ่มส่งเท่านั้น" เหมือนกัน ไม่ยิงระหว่างพิมพ์
# ---------------------------------------------------------------------------
if st.session_state.scenario_submitted:
    followup_text = st.chat_input("Any follow-up question?")

    if followup_text:
        if not api_key:
            st.warning("ยังไม่ได้ตั้งค่า OPENAI_API_KEY จึงเรียก AI ไม่ได้")
        else:
            st.session_state.messages.append({"role": "user", "content": followup_text})
            with st.chat_message("user"):
                st.markdown(followup_text)

            with st.chat_message("assistant"):
                with st.spinner("assessing..."):
                    try:
                        answer = call_openai(st.session_state.messages, model_name)
                    except Exception as e:
                        answer = f"เกิดข้อผิดพลาดในการเรียก OpenAI API: {e}"
                    st.markdown(answer)

            st.session_state.messages.append({"role": "assistant", "content": answer})

    if st.button("Analyze new question."):
        st.session_state.messages = []
        st.session_state.scenario_submitted = False
        st.rerun()

# ---------------------------------------------------------------------------
# ปุ่มดาวน์โหลดบทสนทนาเป็น .docx — อยู่ล่างสุดของหน้า
# สร้างไฟล์ใหม่ทุกครั้งที่ re-render เพื่อให้ได้ประวัติล่าสุด "ณ เวลาที่กดปุ่ม" เสมอ
# ---------------------------------------------------------------------------
if st.session_state.messages:
    st.divider()
    docx_buffer = build_docx(st.session_state.messages)
    st.download_button(
        label="⬇️ Download this discussion",
        data=docx_buffer,
        file_name=f"contract_formation_discussion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )