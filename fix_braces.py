import re

with open("app.py", "r", encoding="utf-8") as f:
    text = f.read()

# We need to replace single `{` and `}` with `{{` and `}}` in the newly added javascript
start_marker = "var ADV_DEF = {"
end_marker = "function buildHdr(tr){"

start_idx = text.find(start_marker)
end_idx = text.find(end_marker, start_idx) + len(end_marker)

if start_idx != -1 and end_idx != -1:
    section = text[start_idx:end_idx]
    
    # Replace ONLY inside this section:
    # We want to replace `{` with `{{` and `}` with `}}`
    # BUT wait, the patch_dashboard already added some python variables maybe? No, the new block is just JS.
    def repl(match):
        ch = match.group(0)
        return "{{" if ch == "{" else "}}"
        
    new_section = re.sub(r'[{}]', repl, section)
    
    new_text = text[:start_idx] + new_section + text[end_idx:]
    with open("app.py", "w", encoding="utf-8") as f:
        f.write(new_text)
    print("Fixed curly braces!")
else:
    print("Could not find markers")
