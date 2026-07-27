#!/usr/bin/env python3
"""
Render LaTeX formulas to PNG images for PPTX embedding.

Usage:
    python render_formula.py <formula_list.json> <output_dir>

Output:
    <output_dir>/formula_XX.png for each complex formula
"""

import json
import os
import sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

def render_latex_to_png(latex_str, output_path, dpi=300, fontsize=14):
    """
    Render a LaTeX formula to a PNG image.
    
    Args:
        latex_str: LaTeX formula string (without the outer $)
        output_path: Path to save the PNG
        dpi: Output resolution (default 300)
        fontsize: Font size for the formula (default 14)
    
    Returns:
        Tuple of (width_inches, height_inches)
    """
    # Estimate figure size based on formula length
    # Simple heuristic: average character width ~0.5em
    estimated_width = len(latex_str) * 0.5 * fontsize / 72 + 0.5
    estimated_height = max(0.8, estimated_width * 0.3)
    
    fig, ax = plt.subplots(figsize=(estimated_width, estimated_height))
    ax.axis('off')
    
    # Center the formula
    ax.text(0.5, 0.5, f'${latex_str}$',
            fontsize=fontsize,
            horizontalalignment='center',
            verticalalignment='center',
            transform=ax.transAxes)
    
    # Save with transparent background
    fig.savefig(output_path, dpi=dpi, bbox_inches='tight',
                pad_inches=0.1, transparent=True)
    plt.close(fig)
    
    return (estimated_width, estimated_height)

def main():
    if len(sys.argv) != 3:
        print(f"Usage: {sys.argv[0]} <formula_list.json> <output_dir>")
        sys.exit(1)
    
    formula_list_path = sys.argv[1]
    output_dir = sys.argv[2]
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Load formula list
    with open(formula_list_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    formulas = data.get('formulas', [])
    if not formulas:
        print("No formulas found in the input file.")
        sys.exit(0)
    
    rendered_count = 0
    skipped_count = 0
    
    for formula in formulas:
        formula_id = formula.get('id', f'formula_{rendered_count+1:02d}')
        latex = formula.get('latex', '')
        complexity = formula.get('complexity', 'simple')
        render_as = formula.get('render_as', 'text')
        
        if render_as != 'png':
            skipped_count += 1
            continue
        
        if not latex:
            print(f"Warning: Formula {formula_id} has empty LaTeX, skipping.")
            skipped_count += 1
            continue
        
        output_path = os.path.join(output_dir, f'{formula_id}.png')
        
        try:
            width, height = render_latex_to_png(latex, output_path)
            formula['rendered_path'] = output_path
            formula['rendered_width'] = width
            formula['rendered_height'] = height
            rendered_count += 1
            print(f"Rendered: {formula_id} -> {output_path}")
        except Exception as e:
            print(f"Error rendering {formula_id}: {e}")
            skipped_count += 1
    
    # Write back the updated formula list with rendered dimensions
    with open(formula_list_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    
    print(f"\nDone: {rendered_count} formulas rendered, {skipped_count} skipped.")

if __name__ == '__main__':
    main()