import fitz  # PyMuPDF
import math

def extract_form_checkboxes_and_labels(page: fitz.Page, x_search_threshold: float = 140.0):
    """
    المرحلة الثانية:
    1. استخراج الأشكال والمربعات الهندسية عبر فحص مسارات الرسم المتجهية (Drawings).
    2. استخراج الكلمات المحيطة بها.
    3. الربط المكاني (Spatial Association) بين كل مربع والوصف التابع له (للنصوص العربية RTL).
    """
    drawings = page.get_drawings()
    checkbox_candidates = []

    # 1. فلترة الأشكال للتعرف على مربعات الاختيار (نسبة العرض للارتفاع والأبعاد القياسية)
    for draw in drawings:
        rect = draw["rect"]
        w = rect.width
        h = rect.height

        # أبعاد مربع الاختيار المعتادة في النماذج (بين 7 و 22 نقطة)
        if 7 <= w <= 22 and 7 <= h <= 22:
            aspect_ratio = w / h
            # التأكد أن الشكل مربع متساوي الأضلاع تقريباً
            if 0.8 <= aspect_ratio <= 1.2:
                # تجنب التكرار لنفس المربع
                if not any(abs(c.x0 - rect.x0) < 2 and abs(c.y0 - rect.y0) < 2 for c in checkbox_candidates):
                    checkbox_candidates.append(rect)

    # 2. استخراج الكلمات مع الإحداثيات: (x0, y0, x1, y1, text, block, line, word_idx)
    words = page.get_text("words")
    
    paired_form_fields = []

    # 3. الربط الهندسي: البحث عن الكلمات التي تقع على يسار المربع في نفس السطر (RTL)
    for box in checkbox_candidates:
        box_center_y = (box.y0 + box.y1) / 2
        matched_words = []

        for w in words:
            wx0, wy0, wx1, wy1, word_text = w[0], w[1], w[2], w[3], w[4]
            word_center_y = (wy0 + wy1) / 2

            # شرط المحاذاة الرأسية (نفس السطر)
            if abs(box_center_y - word_center_y) <= 5.0:
                # شرط الموقع في اللغة العربية: النص يقع على يسار المربع
                distance_x = box.x0 - wx1
                if 0 <= distance_x <= x_search_threshold:
                    matched_words.append((wx0, word_text))

        # فرز الكلمات من اليمين إلى اليسار لبناء الجملة الصحيحة بجانب المربع
        matched_words.sort(key=lambda item: item[0], reverse=True)
        label_text = " ".join([item[1] for item in matched_words])

        paired_form_fields.append({
            "rect": (box.x0, box.y0, box.x1, box.y1),
            "label": label_text if label_text else "بدون وصف",
            "y": box.y0,
            "x": box.x0
        })

    # ترتيب عناصر النموذج من الأعلى للأسفل
    paired_form_fields.sort(key=lambda item: item["y"])
    return paired_form_fields
