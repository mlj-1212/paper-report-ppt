#!/usr/bin/env python3
"""
Intelligent image filtering and sorting for paper-report-ppt.

Reads image_manifest.json and applies filtering, deduplication, 
sorting, and page assignment rules.

Usage:
    python filter_images.py <image_manifest.json> <outline.md> [--max-per-page 2]

Output:
    Updated image_manifest.json with filtered/ranked image list
    Console output: page assignment summary
"""

import json
import os
import sys
import hashlib


# Priority scores for different image types
PRIORITY_MAP = {
    'figure_captioned': 5,       # Figure N: ...
    'fig_captioned': 5,          # Fig. N. ...
    'supplementary_figure': 3,   # Figure S1, Fig S1
    'inline_diagram': 3,         # Small inline diagram
    'chemical_structure': 3,     # Chemical structure
    'author_photo': 1,           # Author photo (filter out)
    'logo': 1,                   # Institution logo (filter out)
    'qr_code': 1,                # QR code (filter out)
    'decorative': 1,             # Decorative illustration (filter out)
}

# Section ordering for sorting
SECTION_ORDER = {
    'introduction': 0,
    'background': 1,
    'methods': 2,
    'results': 3,
    'discussion': 4,
    'conclusion': 5,
    'supplementary': 6,
}


def classify_image(img):
    """Classify an image based on its caption and metadata."""
    caption = (img.get('caption') or '').lower()
    caption_text = (img.get('caption_text') or '').lower()
    source_kind = img.get('source_kind', '')
    
    # Check for figure captions
    if 'figure' in caption_text or 'figure' in caption:
        if 'supplementary' in caption_text or 'supplementary' in caption or \
           'fig s' in caption_text or 'fig s' in caption:
            return 'supplementary_figure'
        return 'figure_captioned'
    
    if caption_text.startswith('fig.') or caption_text.startswith('fig '):
        if 's' in caption_text[:10]:
            return 'supplementary_figure'
        return 'fig_captioned'
    
    # Check for decorative/logos
    small_dim = img.get('pixel_height', 1000) < 100 and img.get('pixel_width', 1000) < 100
    if small_dim and not caption:
        return 'decorative'
    
    # Check for logos
    if 'logo' in caption_text or 'logo' in caption:
        return 'logo'
    
    # Check for QR codes
    if 'qr' in caption_text or 'qr' in caption:
        return 'qr_code'
    
    # Check for chemical structures
    if 'structure' in caption_text or 'chemical' in caption_text:
        return 'chemical_structure'
    
    # Default: inline diagram
    return 'inline_diagram'


def should_keep(img_type):
    """Determine if an image type should be kept."""
    priority = PRIORITY_MAP.get(img_type, 3)
    return priority >= 3


def get_priority(img_type):
    """Get priority score for an image type."""
    return PRIORITY_MAP.get(img_type, 3)


def estimate_section(img):
    """Estimate which section (IMRaD) an image belongs to."""
    caption = (img.get('caption') or '').lower()
    caption_text = (img.get('caption_text') or '').lower()
    page = img.get('page_number', 0)
    
    combined = f"{caption} {caption_text}"
    
    if any(w in combined for w in ['model', 'mechanism', 'work model', 'schematic']):
        return 'discussion'
    if any(w in combined for w in ['method', 'protocol', 'procedure', 'workflow', 'pipeline']):
        return 'methods'
    if any(w in combined for w in ['result', 'data', 'expression', 'response', 'effect']):
        return 'results'
    if any(w in combined for w in ['background', 'overview', 'introduction']):
        return 'introduction'
    if any(w in combined for w in ['conclusion', 'summary']):
        return 'conclusion'
    if any(w in combined for w in ['supplementary', 'figure s', 'fig s']):
        return 'supplementary'
    
    # Default to 'results' for most scientific figures
    return 'results'


def filter_and_sort_images(manifest_path, outline_path=None, max_per_page=2):
    """Main filtering and sorting logic."""
    
    with open(manifest_path, 'r', encoding='utf-8') as f:
        manifest = json.load(f)
    
    images = manifest if isinstance(manifest, list) else manifest.get('images', [])
    
    if not images:
        print("No images found in manifest.")
        return []
    
    # Step 1: Classify and filter
    kept = []
    filtered_out = []
    seen_sha256 = set()
    
    for img in images:
        img_type = classify_image(img)
        img['classified_type'] = img_type
        img['priority'] = get_priority(img_type)
        
        if not should_keep(img_type):
            filtered_out.append(img)
            print(f"  Filtered: {img.get('filename', '?')} (type: {img_type})")
            continue
        
        # Deduplicate by sha256
        sha = img.get('sha256', '')
        if sha and sha in seen_sha256:
            print(f"  Deduped: {img.get('filename', '?')} (duplicate of existing)")
            continue
        if sha:
            seen_sha256.add(sha)
        
        # Estimate section
        img['section'] = estimate_section(img)
        kept.append(img)
    
    print(f"\nKept {len(kept)} images, filtered {len(filtered_out)} images.")
    
    # Step 2: Sort by section order, then by page number
    kept.sort(key=lambda x: (
        SECTION_ORDER.get(x.get('section', 'results'), 99),
        x.get('page_number', 0),
        x.get('figure_number', 0) or 0
    ))
    
    # Step 3: Assign to pages
    section_counts = {}
    for img in kept:
        sec = img.get('section', 'results')
        section_counts[sec] = section_counts.get(sec, 0) + 1
    
    print("\nSection distribution:")
    for sec, count in section_counts.items():
        print(f"  {sec}: {count} images")
    
    # Step 4: Assign page slots
    current_section = None
    page_index = 0
    
    for img in kept:
        sec = img.get('section', 'results')
        if sec != current_section:
            current_section = sec
            page_index = 0
        
        img['page_slot'] = page_index
        page_index += 1
    
    # Step 5: Write back updated manifest
    output = kept if isinstance(manifest, list) else {**manifest, 'images': kept}
    output_path = manifest_path.replace('.json', '_filtered.json')
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output, f, indent=2, ensure_ascii=False)
    
    print(f"\nFiltered manifest written to: {output_path}")
    
    # Print summary
    print("\nImage assignment summary:")
    for img in kept:
        print(f"  [{img.get('section', '?')}] {img.get('filename', '?')} "
              f"(Figure {img.get('figure_number', '?')}, "
              f"Page {img.get('page_number', '?')}, "
              f"Priority {img.get('priority', '?')})")
    
    return kept


def main():
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <image_manifest.json> [outline.md] [--max-per-page N]")
        sys.exit(1)
    
    manifest_path = sys.argv[1]
    outline_path = None
    max_per_page = 2
    
    for arg in sys.argv[2:]:
        if arg == '--max-per-page' and len(sys.argv) > sys.argv.index(arg) + 1:
            max_per_page = int(sys.argv[sys.argv.index(arg) + 1])
        elif arg.endswith('.md'):
            outline_path = arg
    
    if not os.path.exists(manifest_path):
        print(f"Error: {manifest_path} not found.")
        sys.exit(1)
    
    filter_and_sort_images(manifest_path, outline_path, max_per_page)


if __name__ == '__main__':
    main()