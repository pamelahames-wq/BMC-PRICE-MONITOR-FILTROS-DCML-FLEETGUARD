"""merge.py OUT page1 page2 ... : concatena páginas salvas do read_resource (header repetido, rodapé removido)."""
import re, sys
out, pages = sys.argv[1], sys.argv[2:]
foot = re.compile(r"^\[page \d+ .*\]$|^\[.*last page.*\]$")
with open(out, "w", encoding="utf-8") as fo:
    for i, p in enumerate(pages):
        lines = open(p, encoding="utf-8").read().split("\n")
        while lines and (lines[-1].strip() == "" or foot.match(lines[-1].strip())):
            lines.pop()
        if i > 0:
            lines = lines[1:]
        fo.write("\n".join(lines) + "\n")
print(out)
