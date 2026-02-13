# Ketryx Template Generation Agent

A simple AI agent that transforms completed Word documents into Ketryx templates.

## How It Works

1. **You provide:**
   - A finished `.docx` file (the example document)
   - `ketryx_project_data.json` (available fields, types, relations from Ketryx)
   - `ketryx_template_syntax.json` (templating syntax reference)

2. **The agent:**
   - Analyzes the document structure
   - Pattern matches content to available Ketryx data
   - Generates a new `.docx` with template syntax

3. **You get:**
   - A template that matches the original structure exactly
   - Dynamic content replaced with `{project.name}`, `{#items}...{/items}`, etc.

## Installation

```bash
pip install -r requirements.txt
```

You'll also need:
- `pandoc` (for document conversion)
- `node` and `npm` (for docx generation)

```bash
# macOS
brew install pandoc node

# Ubuntu/Debian
sudo apt install pandoc nodejs npm
```

## Usage

```bash
export ANTHROPIC_API_KEY="your-api-key"

python ketryx_template_agent.py \
  --docx "Defect_Summary.docx" \
  --data "ketryx_project_data.json" \
  --syntax "ketryx_template_syntax.json" \
  --output "Defect_Summary_Template.docx"
```

### Options

| Flag | Description | Default |
|------|-------------|---------|
| `--docx` | Input Word document | (required) |
| `--data` | Ketryx project data JSON | (required) |
| `--syntax` | Templating syntax reference JSON | (required) |
| `--output` | Output template path | (required) |
| `--working-dir` | Temp files directory | `./work` |
| `--model` | Claude model | `claude-sonnet-4-20250514` |
| `--max-turns` | Max agent iterations | `50` |

### Using Opus 4.5

For best results on complex documents:

```bash
python ketryx_template_agent.py \
  --model "claude-opus-4-20250514" \
  --docx input.docx \
  --data data.json \
  --syntax syntax.json \
  --output template.docx
```

## What the Agent Does

The agent is given three tools and a system prompt. It figures out the rest:

1. **Reads** the original document (unpacks, converts to readable format)
2. **Reads** the Ketryx data (understands available fields/types)
3. **Reads** the syntax reference (learns how to write templates)
4. **Thinks** about patterns:
   - "These 50 table rows look like defects → loop over KQL results"
   - "This ID format matches `docId` field"
   - "This section references risks → use relation filters"
5. **Generates** a Node.js script using `docx-js`
6. **Runs** the script to create the template

## Example Patterns It Recognizes

| Document Content | Template Replacement |
|------------------|---------------------|
| "Project Name v1.2.0" | `{project.name} {version.name}` |
| Table with 50 defect rows | `{#defects}...{/defects}` loop |
| "DEFECT-123" | `{docId}` |
| "69 issues" | `{defects \| count} issue(s)` |
| Related requirement IDs | `{relations \| where:'type == "implements"' \| map:'other.docId'}` |

## Limitations

- Binary content (images, embedded objects) may need manual handling
- Very complex nested tables might need iteration
- The agent makes educated guesses; review output for accuracy

## Debugging

The agent prints its thinking and tool calls. If something goes wrong:

1. Check the `--working-dir` for intermediate files
2. Look at the generated Node.js script
3. Increase `--max-turns` if it stops too early
4. Try using Opus for more complex reasoning

