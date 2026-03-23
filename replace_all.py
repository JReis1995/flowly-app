import os
import re
import glob

# Search in all .html and .js files
files_to_process = glob.glob(r'c:\Users\joser\OneDrive\Ambiente de Trabalho\Pasta Geral\Flowly\Códigos\Cursor\*.html') + \
                   glob.glob(r'c:\Users\joser\OneDrive\Ambiente de Trabalho\Pasta Geral\Flowly\Códigos\Cursor\*.js')

replacements = {
    r'text-violet-500': 'text-flowly-primary',
    r'text-indigo-500': 'text-flowly-primary',
    r'text-sky-500': 'text-flowly-primary',
    r'bg-red-500': 'bg-flowly-danger',
    r'text-red-600': 'text-flowly-danger',
    r'border-red-200': 'border-flowly-danger/30',
    r'text-amber-600': 'text-flowly-warning',
    r'bg-amber-50': 'bg-flowly-warning/10',
    r'border-slate-200': 'border-flowly-border',
    r'border-slate-300': 'border-flowly-border',
    r'border-\[\#E2E8F0\]': 'border-flowly-border',
    
    # ensure it covers any additional missing ones from the prompt logic
}

for path in files_to_process:
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Apply general replacements
    for old, new in replacements.items():
        content = re.sub(old, new, content)

    # JS charts colors for all JS files
    if path.endswith('.js') or 'JS_' in path:
        content = re.sub(r"'#8b5cf6'", "'#06B6D4'", content) # violet-500
        content = re.sub(r"'#6366f1'", "'#06B6D4'", content) # indigo-500
        content = re.sub(r"'#0ea5e9'", "'#06B6D4'", content) # sky-500
        content = re.sub(r"'#3b82f6'", "'#06B6D4'", content) # blue-500
        content = re.sub(r"'#10b981'", "'#10B981'", content)
        content = re.sub(r"'#22c55e'", "'#10B981'", content) # green-500

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
        
print("Global replacements complete! Processed:", len(files_to_process), "files")
