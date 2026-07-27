import re

with open(r'c:\Users\Homes247\Desktop\office-suite\frontend\src\app\pages\doc-editor\doc-editor.component.ts', 'r', encoding='utf-8') as f:
    content = f.read()

# 1. We want to remove `<div class="menu-bar" ...> ... </div>` from its current place inside `.header-left`.
# 2. We want to place it directly AFTER `<div class="header-bar"> ... </div>`.

# Find menu-bar block
menu_start = content.find('<div class="menu-bar"')
if menu_start == -1:
    print("menu-bar not found")
    exit(1)

# Find the end of menu-bar by matching divs
div_count = 0
menu_end = -1
for i in range(menu_start, len(content)):
    if content.startswith('<div', i):
        div_count += 1
    elif content.startswith('</div', i):
        div_count -= 1
        if div_count == 0:
            menu_end = i + 6
            break

menu_block = content[menu_start:menu_end]

# Remove menu_block from the content
new_content = content[:menu_start] + content[menu_end:]

# Now find where header-bar ends.
header_start = new_content.find('<div class="header-bar">')
div_count = 0
header_end = -1
for i in range(header_start, len(new_content)):
    if new_content.startswith('<div', i):
        div_count += 1
    elif new_content.startswith('</div', i):
        div_count -= 1
        if div_count == 0:
            header_end = i + 6
            break

# Wrap the menu block in a new container
new_menu_block = f"""
      <div class="menu-bar-row" style="display:flex; align-items:center; padding: 2px 12px; background: #fff; border-bottom: 1px solid #dadce0; position: relative; z-index: 205;">
        {menu_block}
      </div>"""

# Insert it after header_bar
final_content = new_content[:header_end] + new_menu_block + new_content[header_end:]

with open(r'c:\Users\Homes247\Desktop\office-suite\frontend\src\app\pages\doc-editor\doc-editor.component.ts', 'w', encoding='utf-8') as f:
    f.write(final_content)

print("HTML Restructured!")
