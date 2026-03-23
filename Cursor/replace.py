import os
import re

files_to_process = [
    'UI_Dashboard.html',
    'JS_Dashboard.html',
    'UI_SaaS_Admin.html',
    'JS_SaaS_Admin.html',
    'UI_CC_Logistica.html',
    'JS_CC_Logistica.html',
    'UI_AI_Export.html',
    'JS_AI_Export.html',
    'Style.html',
    'JS_Globals.html'
]

replacements = {
    # 2. Intervenção nos ficheiros UI
    r'text-violet-500': 'text-flowly-primary',
    r'text-indigo-500': 'text-flowly-primary',
    r'text-sky-500': 'text-flowly-primary',
    r'bg-red-500': 'bg-flowly-danger',
    r'text-red-600': 'text-flowly-danger',
    r'border-red-200': 'border-flowly-danger/30', # giving a bit of opacity for the border so it's not super red if it was red-200
    r'text-amber-600': 'text-flowly-warning',
    r'bg-amber-50': 'bg-flowly-warning/10',
    r'border-slate-200': 'border-flowly-border',
    r'border-slate-300': 'border-flowly-border',
    r'border-\[\#E2E8F0\]': 'border-flowly-border'
}

for filename in files_to_process:
    path = os.path.join(r'c:\Users\joser\OneDrive\Ambiente de Trabalho\Pasta Geral\Flowly\Códigos\Cursor', filename)
    if not os.path.exists(path):
        continue
    
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Apply general replacements
    for old, new in replacements.items():
        content = re.sub(old, new, content)
        
    # Cards: "Garante que todos os cards usam bg-white e bordas border-flowly-border". We already replaced borders. We can replace common light backgrounds like bg-slate-50 in top level divs with bg-white if needed, but it's hard to target "cards" specifically with regex. We'll leave it to the user's explicit mentions and the border replacement which will catch most card borders.

    # Style.html refactoring
    if filename == 'Style.html':
        content = re.sub(r'--primary-color:.*?;', '--primary-color: #06B6D4;', content)
        content = re.sub(r'--danger-red:.*?;', '--danger-red: #EF4444;', content)
        content = re.sub(r'--toast-success:.*?;', '--toast-success: #10B981;', content)
        content = re.sub(r'--toast-error:.*?;', '--toast-error: #EF4444;', content)
        # Padroniza toast-success para bg-flowly-success if they are tailwind classes in JS_Globals
        
    # JS_Dashboard charts colors
    if filename == 'JS_Dashboard.html':
        # Replace common Chart.js hex colors or variable colors
        # Look for things like borderColor: '#...', backgroundColor: '#...' inside the script
        content = re.sub(r"'#8b5cf6'", "'#06B6D4'", content) # violet-500
        content = re.sub(r"'#6366f1'", "'#06B6D4'", content) # indigo-500
        content = re.sub(r"'#0ea5e9'", "'#06B6D4'", content) # sky-500
        content = re.sub(r"'#3b82f6'", "'#06B6D4'", content) # blue-500
        # success color
        content = re.sub(r"'#10b981'", "'#10B981'", content)
        content = re.sub(r"'#22c55e'", "'#10B981'", content) # green-500
        # we will use primary for the main lines and success for the others.
    
    if filename in ['JS_Globals.html', 'Style.html', 'JS_SaaS_Admin.html']:
        content = content.replace('toast-success', 'flowly-success')
        content = content.replace('toast-error', 'flowly-danger')

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
        
print("Replacements complete!")
