import re
import os
import time

XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards"

def unescape_xml(text):
    if not text: return ""
    return text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&quot;", '"').replace("&#39;", "'")

def extract_tag(block, tag):
    match = re.search(f"<{tag}>(.*?)</{tag}>", block, re.IGNORECASE | re.DOTALL)
    return unescape_xml(match.group(1).strip()) if match else ""

start_time = time.time()
files_list = []
for root, dirs, files in os.walk(XML_DIR):
    for f in files:
        if f.endswith('.xml'):
            files_list.append(os.path.join(root, f))

pat_count = 0
for f in files_list:
    try:
        with open(f, 'r') as fh:
            content = fh.read()
            # ONLY SEARCH IF we see '<patient>' tag to be faster!
            if "<patient>" in content or "<patient " in content:
                for pat_block in re.finditer(r"<patient>(.*?)</patient>", content, re.IGNORECASE | re.DOTALL):
                    block = pat_block.group(1)
                    name = extract_tag(block, "fullname")
                    if not name:
                        fname = extract_tag(block, "firstname")
                        sname = extract_tag(block, "surname")
                        name = f"{fname} {sname}".strip()
                    if name:
                        pat_count += 1
    except Exception as e:
        pass
print(f"Final count: {pat_count} in {time.time() - start_time:.2f} seconds")
