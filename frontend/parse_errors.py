import re
content = open('build_output.txt', encoding='utf-8').read()
props = sorted(set(re.findall(r"Property '([^']+)' does not exist", content)))
print('Missing props:')
for p in props:
    print(' ', p)
