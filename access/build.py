# Собирает access/modBuild.bas (Windows-1251, CRLF) из исходников access/src/*.bas (UTF-8).
# Кодировка 1251 нужна для импорта модуля в редактор VBA русской версии Windows/Access.
import glob, os
here = os.path.dirname(os.path.abspath(__file__))
parts = [open(p, encoding="utf-8").read() for p in sorted(glob.glob(os.path.join(here, "src", "*.bas")))]
text = "".join(parts).replace("\r\n", "\n").replace("\n", "\r\n")
with open(os.path.join(here, "modBuild.bas"), "wb") as f:
    f.write(text.encode("cp1251"))
print("modBuild.bas:", text.count("\r\n"), "строк")
