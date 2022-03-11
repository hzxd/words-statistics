import pandas

import docx
from nltk.probability import FreqDist
import os
import csv

# 不统计单词列表
ban_list = []

if __name__ == '__main__':
    text = ""
    for _, _, filenames in os.walk('./docx'):
        for filename in filenames:
            if 'docx' in filename:
                doc = docx.Document('./docx/'+filename)
                for p in doc.paragraphs:
                    text += (p.text+" ")
    # 过滤非英语单词
    text = ''.join(x for x in text if (ord('A') <= ord(x) <= ord('Z')) or (ord('a') <= ord(x) <= ord('z')) or x == ' ')

    words = [x.lower() for x in text.split(" ") if x.strip() != '' and len(x.strip()) > 1 and x not in ban_list]

    fdist = FreqDist(words)
    # 输出结果
    with open('result.csv','w') as f:
        fnames = ['word', 'freq']
        writer = csv.DictWriter(f, fieldnames=fnames)

        writer.writeheader()

        for word,freq in fdist.items():
            writer.writerow({'word': word, 'freq': freq})


