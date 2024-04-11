def ReadTXT(Path):  # 读取TXT文件,返回文本
    TXT = open(Path)  # 返回TXT文件对象
    Txt = TXT.read()  # 读取文本
    TXT.close()  # 关闭TXT文件对象
    return Txt  # 返回文本