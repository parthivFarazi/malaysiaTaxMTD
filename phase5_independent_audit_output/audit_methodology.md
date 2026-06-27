# Phase 5 Independent Audit Methodology

## Method

Tabula and Camelot were not available in the local Python environment. The audit therefore used pdfminer.six directly as the independent extractor.

The independent extraction path differs from the production parser:

- production parser: pdfplumber word extraction plus geometry-derived row/column bands
- audit parser: pdfminer text-line extraction, y-line grouping, and left-to-right token sequence parsing
- no Parsed_Grid or Canonical_Long values are used to parse PDF rows

## Sampling

Random seed: `2018`
KA1-KA10 sampled pages: `13,16,18,19,23,34,36,42,63,67,83,86,101,105,117,122,135,136,150,152,154,165,168,174,176,188,206,208,222,224,233,234,236,244,261,272,274,275,282,300,303,311,314,317,323,328,329,332,345,349`
KA11-KA20 sampled pages: `363,364,371,377,378,389,390,402,408,417,430,439,447,448,454,455,462,469,470,478,484,485,491,499,506,508,525,528,537,539,542,549,565,589,618,620,635,642,651,658,680,683,689,690,693,697,707,712,718,733`

## Match Rules

- salary_from must match exactly
- salary_to must match exactly
- dependent_code must match exactly
- raw deduction value must match exactly
- `-` is compared as a literal dash, not zero
