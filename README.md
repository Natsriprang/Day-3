# Day-3
# 1.create environment
python -m venv .venv
# 2.activate environment
source .venv/bin/activate
# 3.create requirements.txt file
>touch requirements.txt
 streamlit
 openai
 python-dotenv
 pypdf: ใช้อ่าน text จาก pdf
 chromadb: ใช้อัพ chroma
# 4.install dependencies and freeze version so i can rely on them
>pip install -r requirements.txt
>pip freeze > requirements.txt
freeze ทุกครั้งที่เพิ่มอะไรเข้าไป
# 5.create .env for secrets
>touch .env
>open ai key, file password
# 6.add pages folder
> mkdir pages 
or
create folders
# 7.add streamlit entry point (หน้างานหลัก)

# 8.run streamlit