import re

with open('style.css', 'r', encoding='utf-8') as f:
    css = f.read()

# Replace variables
css = re.sub(
    r':root \{.*?\n\}',
    ''':root {
  --bg: #09090b;
  --panel: rgba(24, 24, 27, 0.75);
  --panel-strong: rgba(39, 39, 42, 0.85);
  --text: #f8fafc;
  --muted: #94a3b8;
  --accent: #8b5cf6;
  --accent-dark: #7c3aed;
  --line: rgba(255, 255, 255, 0.1);
  --shadow: 0 24px 80px rgba(0, 0, 0, 0.6);
}''',
    css,
    flags=re.DOTALL
)

# Fix body background
css = re.sub(
    r'background:\s*radial-gradient.*?linear-gradient.*?;',
    '''background-color: var(--bg);
  background-image: 
    radial-gradient(circle at 15% 50%, rgba(139, 92, 246, 0.15), transparent 25%),
    radial-gradient(circle at 85% 30%, rgba(6, 182, 212, 0.15), transparent 25%);
  background-attachment: fixed;''',
    css,
    flags=re.DOTALL
)

# Replace hardcoded light mode colors with dark mode equivalents
replacements = [
    (r'rgba\(255, 255, 255, 0\.55\)', r'rgba(255, 255, 255, 0.1)'),
    (r'rgba\(255, 255, 255, 0\.75\)', r'rgba(255, 255, 255, 0.05)'),
    (r'rgba\(255, 244, 233, 0\.92\)', r'rgba(24, 24, 27, 0.8)'),
    (r'rgba\(31, 41, 55, 0\.06\)', r'rgba(255, 255, 255, 0.05)'),
    (r'rgba\(31, 41, 55, 0\.08\)', r'rgba(255, 255, 255, 0.08)'),
    (r'rgba\(31, 41, 55, 0\.18\)', r'rgba(255, 255, 255, 0.18)'),
    (r'rgba\(0, 0, 0, 0\.08\)', r'rgba(255, 255, 255, 0.08)'),
    (r'rgba\(255, 255, 255, 0\.72\)', r'rgba(24, 24, 27, 0.72)'),
    (r'rgba\(255, 255, 255, 0\.65\)', r'rgba(24, 24, 27, 0.65)'),
    (r'rgba\(255, 255, 255, 0\.45\)', r'rgba(24, 24, 27, 0.45)'),
    (r'rgba\(255, 255, 255, 0\.68\)', r'rgba(24, 24, 27, 0.68)'),
    (r'rgba\(194, 74, 26, 0\.18\)', r'rgba(244, 63, 94, 0.18)'),
    (r'rgba\(239, 108, 51, 0\.1\)', r'rgba(244, 63, 94, 0.1)'),
    (r'rgba\(239, 108, 51, 0\.22\)', r'rgba(139, 92, 246, 0.22)'),
    (r'rgba\(239, 108, 51, 0\.35\)', r'rgba(139, 92, 246, 0.35)'),
    (r'rgba\(239, 108, 51, 0\.45\)', r'rgba(139, 92, 246, 0.45)'),
    (r'rgba\(194, 74, 26, 0\.28\)', r'rgba(244, 63, 94, 0.28)'),
    (r'rgba\(255, 248, 239, 0\.92\)', r'rgba(24, 24, 27, 0.92)'),
    (r'background: #fff7ed;', r'background: rgba(39, 39, 42, 0.9);'),
    (r'background: #fffaf3;', r'background: var(--panel-strong);'),
    (r'box-shadow: 0 30px 80px rgba\(15, 23, 42, 0\.28\);', r'box-shadow: 0 30px 80px rgba(0, 0, 0, 0.8);')
]

for old, new in replacements:
    css = re.sub(old, new, css)

with open('style.css', 'w', encoding='utf-8') as f:
    f.write(css)
print("Updated style.css successfully!")
