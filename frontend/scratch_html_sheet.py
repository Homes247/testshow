import re

with open(r'c:\Users\Homes247\Desktop\office-suite\frontend\src\app\pages\sheet-editor\sheet-editor.component.ts', 'r', encoding='utf-8') as f:
    content = f.read()

menu_start = content.find('<div class="menu-bar"')
if menu_start == -1:
    print("menu-bar not found in sheet")
    exit(0)

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
new_content = content[:menu_start] + content[menu_end:]

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

new_menu_block = f"""
      <div class="menu-bar-row" style="display:flex; align-items:center; padding: 2px 12px; background: #fff; border-bottom: 1px solid #dadce0; position: relative; z-index: 205;">
        {menu_block}
      </div>"""

final_content = new_content[:header_end] + new_menu_block + new_content[header_end:]

with open(r'c:\Users\Homes247\Desktop\office-suite\frontend\src\app\pages\sheet-editor\sheet-editor.component.ts', 'w', encoding='utf-8') as f:
    f.write(final_content)

print("Sheet HTML Restructured!")
