#!/usr/bin/env python3
"""
Script to organize extracted VBA source code from olevba output into structured folders.

Takes the output files from olevba (e.g., A_vba_source.txt) and creates:
- A folder for each workbook (e.g., A_vba/, Ab_vba/, etc.)
- Individual files for each VBA module/class within those folders
"""

import os
import re
from pathlib import Path


def parse_vba_source_file(source_file_path):
    """Parse a VBA source file and extract individual modules."""
    with open(source_file_path, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    modules = []
    lines = content.split('\n')
    i = 0
    
    while i < len(lines):
        line = lines[i].strip()
        
        # Look for VBA MACRO declarations
        if line.startswith('VBA MACRO '):
            # Extract module name
            module_name = line.replace('VBA MACRO ', '').strip()
            
            # Skip the next line (in file: ...)
            i += 1
            
            # Skip the separator line (- - - -)
            i += 1
            if i < len(lines) and lines[i].strip().startswith('- - - -'):
                i += 1
            
            # Collect module content until next separator or end
            module_content = []
            while i < len(lines):
                line = lines[i]
                
                # Stop at next module separator
                if line.strip().startswith('-------------------------------------------------------------------------------'):
                    break
                    
                # Stop at next VBA MACRO declaration
                if line.strip().startswith('VBA MACRO '):
                    i -= 1  # Back up one line so outer loop processes this
                    break
                
                module_content.append(line)
                i += 1
            
            # Store the module
            content_str = '\n'.join(module_content).strip()
            if content_str and content_str != '(empty macro)':
                modules.append({
                    'name': module_name,
                    'content': content_str
                })
        
        i += 1
    
    return modules


def sanitize_filename(filename):
    """Sanitize filename to be safe for filesystem."""
    # Replace invalid characters
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
    return filename


def create_vba_folders(source_files):
    """Create organized folder structure for all VBA source files."""
    
    for source_file in source_files:
        if not source_file.endswith('_vba_source.txt'):
            continue
            
        # Extract workbook name (e.g., "A" from "A_vba_source.txt")
        workbook_name = source_file.replace('_vba_source.txt', '')
        
        print(f"Processing {source_file} -> {workbook_name}_vba/")
        
        # Parse the VBA source file
        modules = parse_vba_source_file(source_file)
        
        if not modules:
            print(f"  No modules found in {source_file}")
            continue
            
        # Create output folder
        output_folder = Path(f"{workbook_name}_vba")
        output_folder.mkdir(exist_ok=True)
        
        # Create individual files for each module
        for module in modules:
            module_name = sanitize_filename(module['name'])
            
            # Determine file extension based on module type
            if module_name.endswith('.cls'):
                # Class module
                output_file = output_folder / module_name
            elif module_name.endswith('.bas'):
                # Standard module
                output_file = output_folder / module_name
            elif module_name.endswith('.frm'):
                # Form module
                output_file = output_folder / module_name
            else:
                # Default to .bas if no extension
                output_file = output_folder / f"{module_name}.bas"
            
            # Skip empty modules
            content = module['content']
            if not content or content == '(empty macro)':
                print(f"  Skipping empty module: {module_name}")
                continue
                
            # Write module content to file
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"  Created: {output_file}")
        
        print(f"  Completed {workbook_name}_vba/ with {len([m for m in modules if m['content'] and m['content'] != '(empty macro)'])} modules")
        print()


def main():
    """Main function to organize all VBA source files."""
    # Find all VBA source files in current directory
    source_files = [f for f in os.listdir('.') if f.endswith('_vba_source.txt')]
    
    if not source_files:
        print("No VBA source files found (*_vba_source.txt)")
        return
    
    print(f"Found {len(source_files)} VBA source files:")
    for f in source_files:
        print(f"  - {f}")
    print()
    
    create_vba_folders(source_files)
    
    print("VBA source organization complete!")


if __name__ == "__main__":
    main()