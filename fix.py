import re

with open("app.py", "r") as f:
    content = f.read()

old = '''            else:
                cell_date = datetime.date.fromisoformat(str(cell_val))'''

new = '''            else:
                cell_val_str = str(cell_val).replace("/", "-")
                cell_date = datetime.date.fromisoformat(cell_val_str)'''

content = content.replace(old, new)

with open("app.py", "w") as f:
    f.write(content)

print("修正完了！")
