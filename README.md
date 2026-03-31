# Sahaj Presentation Plugin for Claude Code

A [Claude Code](https://docs.anthropic.com/en/docs/claude-code) plugin that generates professional PowerPoint presentations in Sahaj's corporate visual identity.

## What it does

- Reads reference documents and structures content into slides
- Generates `.pptx` files with Sahaj branding (colors, fonts, layouts)
- Supports title slides, section dividers, bullet content, and card layouts
- Two-phase workflow: content structuring (with user review) then generation

## Installation

### Option 1: Via plugin marketplace (recommended)

If this plugin is listed in a marketplace you have access to:

```
/plugin install sahaj-presentation
```

### Option 2: Local install from GitHub

```bash
git clone https://github.com/priyadarshan85/sahaj-presentation-skill.git
```

Then start Claude Code with:

```bash
claude --plugin-dir ./sahaj-presentation-skill
```

Or to install permanently, copy it into your plugins directory:

```bash
cp -r sahaj-presentation-skill ~/.claude/plugins/sahaj-presentation
```

Python dependencies (`python-pptx`, `lxml`) are installed automatically on first session via the plugin's setup hook.

## Usage

In Claude Code, just ask to create a presentation:

```
Create a presentation from this document: [path/to/document]
```

Or use the slash command:

```
/sahaj-presentation:sahaj-presentation
```

The skill will:
1. Read your reference content
2. Propose a slide-by-slide structure for your review
3. Generate the `.pptx` after you confirm

## Slide Types

| Type | Use for |
|------|---------|
| **Title** | Opening slide with Sahaj logo + optional client logo |
| **Section Divider** | Dark blue slide between major sections |
| **Bullet Content** | Narrative, sequential steps, detailed points |
| **Card Content** | Parallel items, comparisons, metrics side-by-side |

## Brand Identity

- **Heading font**: Zilla Slab
- **Body font**: Mulish
- **Primary color**: #002060 (dark blue)
- **Card heading color**: #6061AD (purple)
- **Slide size**: 16:9 widescreen (10.0" x 5.625")

## Requirements

- [Claude Code](https://docs.anthropic.com/en/docs/claude-code)
- Python 3.9+
- `python-pptx` and `lxml` (auto-installed by plugin hook)

## Plugin Structure

```
sahaj-presentation-skill/
├── .claude-plugin/
│   └── plugin.json           # Plugin manifest
├── skills/
│   └── sahaj-presentation/
│       ├── SKILL.md           # Skill definition
│       ├── assets/            # Logos and images
│       └── scripts/           # Python generator
├── hooks/
│   └── hooks.json            # Auto-setup hook
├── README.md
└── .gitignore
```
