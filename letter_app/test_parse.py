import re
import os
import sys

XML_DIR = "/Users/cbasnayake/Documents/BACKUP CMG XML/2017 onwards"

def unescape_xml(text):
    if not text: return ""
    return text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&").replace("&quot;", '"').replace("&#39;", "'")

def extract_tag(block, tag):
    match = re.search(f"<{tag}>(.*?)</{tag}>", block, re.IGNORECASE | re.DOTALL)
    return unescape_xml(match.group(1).strip()) if match else ""

pat_count = 0
files_list = []
for root, dirs, files in os.walk(XML_DIR):
    for f in files:
        if f.endswith('.xml'):
            files_list.append(os.path.join(root, f))
print(f"Total XMLs found: {len(files_list)}")

if len(files_list) > 0:
    for i in range(10): # test first 10
        with open(files_list[i], 'r') as fh:
            content = fh.read()
            for pat_block in re.finditer(r"<patient>(.*?)</patient>", content, re.IGNORECASE | re.DOTALL):
                block = pat_block.group(1)
                name = extract_tag(block, "fullname")
                print("Found match:", name)
                pat_count += 1
print("Final count in 10 files:", pat_count)
