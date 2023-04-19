import collections.abc
from pptx import Presentation


# get paragraph array by each slide
def get_text_data(presentation):
    slide_num = len(presentation.slides)
    text_db = []
    for i in range(slide_num):
        slide = prs.slides[i]
        # print(slide)
        texts = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text_frame = shape.text_frame
            for paragraph in text_frame.paragraphs:
                text = paragraph.text
                # print(text)
                if(text):
                    texts.append(text)
        if(True):
            text_db.append(texts)
    return text_db


#-----------------------------------------main-----------------------------------

prs = Presentation("../data/introduction.pptx")
db = get_text_data(prs)
print(db)

keywords = ["データ", "ビジュアライズ", "ビジュアライ","プログラミング"]
    
result = []
for paragraphs in db:
    count_per_key = []
    for key in keywords:
        key_count = 0
        for paragraph in paragraphs:
            count = paragraph.count(key)
            key_count += count
        count_per_key.append(key_count)
    result.append(count_per_key)

print(result)