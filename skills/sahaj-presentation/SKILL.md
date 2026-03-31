---
name: sahaj-presentation
description: "Create professional PowerPoint presentations in Sahaj's corporate theme and style. Use this skill whenever the user wants to create a presentation, slide deck, pptx, or pitch deck from a document, notes, or content — especially when they mention Sahaj branding, the Sahaj template, or want content organized into slides. Also trigger when the user provides a reference document and asks to 'present this', 'make slides from this', 'turn this into a deck', or similar. This skill handles content structuring, slide layout decisions, and .pptx generation using Sahaj's exact brand colors, fonts, and layout patterns."
model: claude-opus-4-6
---

# Sahaj Presentation Skill

Create polished PowerPoint presentations in Sahaj's corporate visual identity from reference content documents.

## Workflow

The process has two phases: **Structure** (organize the content, get user sign-off) then **Generate** (produce the .pptx).

### Phase 1: Content Structuring

1. Read the user's reference document thoroughly
2. Identify the key themes, sections, and data points
3. Propose a slide-by-slide outline to the user, structured like this:

```
Proposed Presentation Structure:

1. Title Slide — [presentation title] | [client name / context]
2. Section Divider — "Executive Summary"
3. Content Slide — Context: [2-3 bullet summary of what this slide covers]
4. Content Slide (Card Layout) — Problem: [card titles]
5. Content Slide — Proposed Solution: [bullet summary]
6. Content Slide (Card Layout) — Business Impact: [card titles]
7. Content Slide (Card Layout) — Investment & ROI: [card titles]
8. Section Divider — "Recommendation"
9. Content Slide — Next Steps: [bullet summary]
```

4. Wait for the user to confirm or adjust the structure before generating

**Content structuring principles:**
- Lead with context/problem before solution — the audience needs to understand "why" before "what"
- One idea per slide. If a slide has more than 5-6 bullets or 4+ cards, split it
- Use card layouts for parallel concepts (e.g., comparing risks, listing benefits, showing metrics side by side). Use bullet layouts for sequential narrative or explanatory content
- Keep bullet text concise — aim for one line per bullet, two lines max. The slides are a visual aid, not the full document
- Bold the key phrase in each bullet when there's a clear takeaway
- Section dividers go before each major topic shift — they give the audience a mental reset

### Phase 2: Presentation Generation

Once the user confirms the structure, generate the .pptx using the bundled script.

Run the generation script:
```bash
<skill-path>/scripts/generate_presentation.py
```

The script expects a JSON specification on stdin. Build the JSON from the confirmed structure:

```json
{
  "output_path": "Presentation.pptx",
  "client_logo_path": null,
  "slides": [
    {
      "type": "title",
      "title": "Presentation Title Here",
      "client_logo_path": "/path/to/client_logo.png"
    },
    {
      "type": "section_divider",
      "title": "Executive Summary"
    },
    {
      "type": "bullet_content",
      "title": "Context",
      "subtitle": "Brief subtitle here",
      "bullets": [
        {"text": "First point with **bold emphasis** on key phrase", "level": 0},
        {"text": "Sub-point elaborating on the above", "level": 1},
        {"text": "Second main point", "level": 0}
      ]
    },
    {
      "type": "card_content",
      "title": "Problem",
      "subtitle": "Challenges Overview",
      "cards": [
        {
          "heading": "Card Title",
          "body": "Description text for this card. Keep it to 2-3 short lines."
        },
        {
          "heading": "Another Card",
          "body": "Another description. Cards work best in groups of 2-3 per row."
        }
      ]
    }
  ]
}
```

Write this JSON to a temp file and pipe it to the script:

```bash
cat /tmp/presentation_spec.json | <skill-path>/scripts/.venv/bin/python <skill-path>/scripts/generate_presentation.py
```

If the script's venv doesn't exist yet, create it first:
```bash
cd <skill-path>/scripts && python3 -m venv .venv && .venv/bin/pip install python-pptx lxml
```

### Slide Types Reference

**Title Slide** (`type: "title"`)
- Large presentation title centered-left
- Sahaj logo + vertical separator + client logo (if provided)
- White text on white background

**Section Divider** (`type: "section_divider"`)
- Dark blue (#002060) background
- Large white centered title in Zilla Slab bold
- Use between major sections to create visual breathing room

**Bullet Content** (`type: "bullet_content"`)
- Title (Zilla Slab 24pt bold, dark blue) + subtitle (Zilla Slab, dark blue) at top
- Body area with hierarchical bullets:
  - Level 0: filled circle (●) — main points
  - Level 1: hollow circle (○) — sub-points
- Font: Mulish 12pt, color #0E0E0E
- Use `**text**` in bullet text to make portions bold

**Card Content** (`type: "card_content"`)
- Same title/subtitle header as bullet content
- 2-3 cards per row arranged in a grid
- Each card: purple (#6061AD) bold heading in Mulish 11pt, body text in Mulish 11pt #0E0E0E
- Cards auto-arrange: 2 cards = 2 columns, 3 cards = 3 columns, 4-6 cards = 2 rows

### Choosing Between Slide Types

| Content pattern | Slide type |
|---|---|
| Narrative explanation, sequential steps, detailed description | `bullet_content` |
| Parallel categories, comparison items, metrics side-by-side | `card_content` |
| 1-2 sentences max, high-level category label | `section_divider` |

## Sahaj Brand Identity (for reference)

These values are baked into the generation script — you don't need to specify them manually. They're documented here so you understand the visual system:

- **Heading font**: Zilla Slab (bold for titles, regular for subtitles)
- **Body font**: Mulish (regular for body, bold for card headings and emphasis)
- **Title color**: #002060 (dark blue)
- **Body text color**: #0E0E0E (near-black)
- **Card heading color**: #6061AD (purple)
- **Section divider background**: #002060 (dark blue)
- **Accent/link color**: #0C5395 (blue)
- **Slide size**: 10.0" x 5.625" (16:9 widescreen)
- **Bullet styles**: ● for level 0, ○ for level 1, - for level 2

## Important Notes

- Ask the user for a client logo path if they want one on the title slide. If they don't provide one, the title slide will only show the Sahaj logo.
- The generation script is self-contained — it handles all font embedding, color application, and layout math.
- If `python-pptx` isn't installed in the script's venv, set it up first (see command above).
- The output .pptx will be saved to the path specified in the JSON spec. Default to the current working directory with a descriptive filename.
