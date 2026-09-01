import os

filepath = r"r:\DLK-Scan-PDF\Scan-PDF\templates\converter\base.html"

with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

# 1. Fix CSS padding for navigation buttons to reduce gap
css_target_1 = """        #navbar .desktop-nav .group>button {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            line-height: 1;
            padding-left: 8px;
            padding-right: 8px;
            font-size: 12px;
            letter-spacing: 0.05em;
        }"""
css_replacement_1 = """        #navbar .desktop-nav .group>button {
            display: inline-flex;
            align-items: center;
            gap: 4px;
            line-height: 1;
            padding-left: 6px;
            padding-right: 6px;
            font-size: 12px;
            letter-spacing: 0.05em;
        }"""
content = content.replace(css_target_1, css_replacement_1)

css_target_2 = """        @media (min-width: 1280px) {
            #navbar .desktop-nav .group>button {
                gap: 6px;
                padding-left: 16px;
                padding-right: 16px;
                font-size: 14px;
                letter-spacing: 0.1em;
            }
        }"""
css_replacement_2 = """        @media (min-width: 1280px) {
            #navbar .desktop-nav .group>button {
                gap: 6px;
                padding-left: 10px;
                padding-right: 10px;
                font-size: 14px;
                letter-spacing: 0.1em;
            }
        }"""
content = content.replace(css_target_2, css_replacement_2)

# 2. Extract Link Tools block
start_marker = "<!-- Link Tools Menu -->"
end_marker = "<!-- Search Bar: Properly Aligned Right -->"
link_tools_idx = content.find(start_marker)
search_idx = content.find(end_marker)

if link_tools_idx != -1 and search_idx != -1:
    link_tools_block = content[link_tools_idx:search_idx]
    # Remove hidden lg:block to match others
    link_tools_block_clean = link_tools_block.replace('<div class="relative group hidden lg:block">', '<div class="relative group">')
    
    # Remove from old position
    content = content[:link_tools_idx] + content[search_idx:]
    
    # Insert before Services
    services_marker = "<!-- Important Tools (Styled like Image Tools mega) -->"
    services_idx = content.find(services_marker)
    
    if services_idx != -1:
        content = content[:services_idx] + link_tools_block_clean + content[services_idx:]

with open(filepath, "w", encoding="utf-8") as f:
    f.write(content)
