# -*- coding: utf-8 -*-
import google.generativeai as genai
from config import GEMINI_API_KEY, GEMINI_MODEL

genai.configure(api_key=GEMINI_API_KEY)
model = genai.GenerativeModel(GEMINI_MODEL)

def ask(prompt: str) -> str:
    response = model.generate_content(prompt)
    return response.text

if __name__ == "__main__":
    print(ask("안녕하세요! 간단히 자기소개 해줘."))
