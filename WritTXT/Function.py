def WritTXT(Path, Txt):  # 写TXT文件
    TXT = open(Path, 'a', encoding='utf-8')  # 返回文本文件对象
    TXT.write(Txt)
    TXT.close()