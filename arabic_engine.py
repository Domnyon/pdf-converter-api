import fitz  # PyMuPDF
import unicodedata
import re

def is_arabic_char(char: str) -> bool:
    """التحقق مما إذا كان الرمز ينتمي لمجموعات اليونيكود للغة العربية"""
    code = ord(char)
    return (
        (0x0600 <= code <= 0x06FF) or  # النطاق الأساسي للعربية
        (0x0750 <= code <= 0x077F) or  # الملحق العربي
        (0xFB50 <= code <= 0xFDFF) or  # أشكال العرض التقديمي أ
        (0xFE70 <= code <= 0xFEFF)     # أشكال العرض التقديمي ب (الحروف المتصلة المرمزة)
    )

def is_arabic_text(text: str) -> bool:
    """كشف احتواء النص على أحرف عربية"""
    return any(is_arabic_char(c) for c in text)

def fix_visual_arabic_word(word: str) -> str:
    """
    عكس ترتيب الحروف إذا خُزنت بصرياً من اليسار لليمين داخل الـ PDF
    مع الحفاظ على الأرقام والرموز اللاتينية
    """
    if not is_arabic_text(word):
        return word
    return word[::-1]

def extract_and_sort_arabic_lines(page: fitz.Page, y_tolerance: float = 3.5):
    """
    استخراج الكلمات وترتيبها مكانياً:
    1. تجميع الكلمات التي تقع على نفس السطر الرأسي Y.
    2. فرز الكلمات العربية من اليمين إلى اليسار (تنازلياً حسب X).
    """
    words = page.get_text("words")
    if not words:
        return []

    lines = []
    sorted_words = sorted(words, key=lambda w: (w[1], w[0]))

    current_line = []
    current_y = None

    for w in sorted_words:
        x0, y0, x1, y1, text = w[0], w[1], w[2], w[3], w[4]
        word_center_y = (y0 + y1) / 2

        if current_y is None:
            current_y = word_center_y
            current_line.append(w)
        else:
            if abs(word_center_y - current_y) <= y_tolerance:
                current_line.append(w)
            else:
                lines.append((current_y, current_line))
                current_line = [w]
                current_y = word_center_y

    if current_line:
        lines.append((current_y, current_line))

    lines.sort(key=lambda item: item[0])
    processed_lines = []

    for _, line_words in lines:
        has_arabic = any(is_arabic_text(w[4]) for w in line_words)

        if has_arabic:
            line_words.sort(key=lambda w: w[2], reverse=True)
        else:
            line_words.sort(key=lambda w: w[0])

        formatted_line = []
        for w in line_words:
            raw_text = w[4]
            if is_arabic_text(raw_text):
                corrected_word = fix_visual_arabic_word(raw_text)
                formatted_line.append(corrected_word)
            else:
                formatted_line.append(raw_text)

        line_str = " ".join(formatted_line)
        processed_lines.append(line_str)

    return processed_lines
