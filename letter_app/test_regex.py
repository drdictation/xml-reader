import re
texts = [
    "<b>•</b> Reassurance",
    "<i>· Impression</i>",
    "•  Implementation",
    "<b>1.</b> Recommendation",
    "<b>2)</b> Advice",
    "Normal text"
]
for t in texts:
    m = re.match(r'^((?:<[^>]+>|\s)*)([•·\uf0b7●◦])((?:<[^>]+>|\s)*)(.*)$', t)
    if m:
        print("BULLET:", m.group(1) + m.group(3) + m.group(4))
    else:
        m2 = re.match(r'^((?:<[^>]+>|\s)*)(\d+)[.)]((?:<[^>]+>|\s)*)(.*)$', t)
        if m2:
            print("NUMBERED:", m2.group(2) + ".", m2.group(1) + m2.group(3) + m2.group(4))
        else:
            print("NORMAL:", t)
