def WritSCM(Path, Txt):  # 写SCM文件
    from WritTXT.Function import WritTXT
    WritTXT(Path, 'a', Txt)  # 写TXT文件