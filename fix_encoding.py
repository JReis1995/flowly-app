#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script to fix UTF-8 encoding corruption in project files.
Fixes common Latin1/UTF-8 encoding issues.
"""

import os
import sys
import glob

# Define corruption patterns to fix
REPLACEMENTS = [
    ('NÁO', 'NÃO'),  # Uppercase corruption
    ('nÁo', 'não'),  # Mixed case corruption
    ('Áo', 'ão'),
    ('çÁ', 'çã'),
    ('Á ', 'ã '),
    ('ÁÁ', 'ãã'),
    ('Áe', 'ãe'),
    ('Ás', 'ãs'),
    ('Áa', 'ãa'),
    ('Ái', 'ãi'),
    ('Áu', 'ãu'),
    ('Ác', 'ãc'),
    ('Ád', 'ãd'),
    ('Ág', 'ãg'),
    ('Ám', 'ãm'),
    ('Án', 'ãn'),
    ('Áp', 'ãp'),
    ('Ár', 'ãr'),
    ('Át', 'ãt'),
    ('Áv', 'ãv'),
    ('Áb', 'ãb'),
    ('Áf', 'ãf'),
    ('Áh', 'ãh'),
    ('Ál', 'ãl'),
    ('Áq', 'ãq'),
    ('Áw', 'ãw'),
    ('Áx', 'ãx'),
    ('Áy', 'ãy'),
    ('Áz', 'ãz'),
    ('Á.', 'ã.'),
    ('Á,', 'ã,'),
    ('Á;', 'ã;'),
    ('Á:', 'ã:'),
    ('Á)', 'ã)'),
    ('Á}', 'ã}'),
    ('Á]', 'ã]'),
    ('Á"', 'ã"'),
    ("Á'", "ã'"),
    ('Á<', 'ã<'),
    ('Á>', 'ã>'),
    ('Á/', 'ã/'),
    ('Á\\', 'ã\\'),
    ('Á-', 'ã-'),
    ('Á_', 'ã_'),
    ('Á=', 'ã='),
    ('Á+', 'ã+'),
    ('Á*', 'ã*'),
    ('Á&', 'ã&'),
    ('Á%', 'ã%'),
    ('Á$', 'ã$'),
    ('Á#', 'ã#'),
    ('Á@', 'ã@'),
    ('Á!', 'ã!'),
    ('Á?', 'ã?'),
    ('Á|', 'ã|'),
    ('Á\n', 'ã\n'),
    ('Á\r', 'ã\r'),
    ('Á\t', 'ã\t'),
]

# File extensions to process
EXTENSIONS = ['.js', '.gs', '.html', '.css']

def fix_file_encoding(filepath):
    """Fix encoding corruption in a single file."""
    try:
        # Try multiple encodings to read the file
        content = None
        encodings = ['utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
        
        for encoding in encodings:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    content = f.read()
                break
            except (UnicodeDecodeError, UnicodeError):
                continue
        
        if content is None:
            # Last resort: read as binary and decode with errors='replace'
            with open(filepath, 'rb') as f:
                content = f.read().decode('utf-8', errors='replace')
        
        original_content = content
        changes_made = 0
        
        # Apply all replacements
        for old, new in REPLACEMENTS:
            if old in content:
                count = content.count(old)
                content = content.replace(old, new)
                changes_made += count
        
        # Write back only if changes were made
        if content != original_content:
            with open(filepath, 'w', encoding='utf-8', newline='') as f:
                f.write(content)
            return changes_made
        
        return 0
    
    except Exception as e:
        print(f"ERROR processing {filepath}: {e}")
        return 0

def main():
    """Main function to process all files."""
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    print("=" * 60)
    print("Flowly Encoding Fix Script")
    print("=" * 60)
    print(f"Working directory: {current_dir}")
    print(f"Extensions to process: {', '.join(EXTENSIONS)}")
    print("=" * 60)
    
    total_files = 0
    total_changes = 0
    files_modified = 0
    
    # Process each extension
    for ext in EXTENSIONS:
        pattern = os.path.join(current_dir, f'*{ext}')
        files = glob.glob(pattern)
        
        for filepath in files:
            filename = os.path.basename(filepath)
            changes = fix_file_encoding(filepath)
            
            if changes > 0:
                print(f"✓ {filename}: {changes} corrections")
                files_modified += 1
                total_changes += changes
            
            total_files += 1
    
    print("=" * 60)
    print(f"SUMMARY:")
    print(f"  Files scanned: {total_files}")
    print(f"  Files modified: {files_modified}")
    print(f"  Total corrections: {total_changes}")
    print("=" * 60)
    
    if total_changes > 0:
        print("✓ Encoding corruption fixed successfully!")
    else:
        print("No corruption patterns found.")

if __name__ == '__main__':
    main()
