import base64
import os

hero_path = 'hero.png'
icon_path = 'icon.png'

# Check if images exist, if not try to get from artifacts paths provided in history
if not os.path.exists(hero_path):
    print("Hero image missing, cannot recreate assets.py")
    exit(1)

with open(hero_path, 'rb') as f:
    hero_b64 = base64.b64encode(f.read()).decode()

with open(icon_path, 'rb') as f:
    icon_b64 = base64.b64encode(f.read()).decode()

with open('assets.py', 'w', encoding='utf-8') as f:
    f.write(f'HERO_B64 = "{hero_b64}"\n')
    f.write(f'ICON_B64 = "{icon_b64}"\n')
print("Assets converted successfully.")
