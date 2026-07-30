# Image Selection Rules

## Overview

`filter_images.py` applies intelligent filtering to images extracted from PDFs to select only those relevant for presentation slides.

## Filtering Criteria

### 1. Figure Detection (Priority)

Images with captions containing figure identifiers are highest priority:
- `Figure N:` / `Fig. N:` / `Fig N.`
- `Figure S1:` (supplementary figures, lower priority)
- `Scheme N:` (chemical schemes)

### 2. Size Filtering

| Property | Min Threshold | Action |
|---|---|---|
| Width | 100 px | Discard |
| Height | 100 px | Discard |
| Area | 20,000 px² | Discard |

### 3. Content Filtering

Images classified as the following are excluded:
- Institution logos
- Author photos
- QR codes
- Decorative borders/backgrounds
- Small inline icons (< 50×50 px)

### 4. Deduplication

Images with identical SHA256 hashes are deduplicated, keeping only the first occurrence.

## Sorting Rules

After filtering, images are sorted by:
1. **Section order**: Introduction → Methods → Results → Discussion
2. **Figure number**: Figure 1, Figure 2, ...
3. **Page order**: Earlier pages first

## Output

`image_manifest_filtered.json` contains the filtered and sorted image list with:
- `filename`, `sha256`, `page_number`
- `assigned_section`: inferred from caption text
- `priority_score`: 1-5 scale
- `filter_reason`: if excluded
