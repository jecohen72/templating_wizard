#!/usr/bin/env python3
"""
Ketryx Template Generation Agent v6 - File-Based Context
"""

import anthropic
import argparse
import base64
import json
import sys
from pathlib import Path


PRICING = {
    "claude-sonnet-4-5-20250929": {"input": 3.0, "output": 15.0},
    "claude-opus-4-5-20250514": {"input": 15.0, "output": 75.0},
}


def estimate_cost(model: str, input_tokens: int, output_tokens: int) -> float:
    p = PRICING.get(model, PRICING["claude-sonnet-4-5-20250929"])
    return (input_tokens / 1_000_000) * p["input"] + (output_tokens / 1_000_000) * p["output"]


def upload_file(client: anthropic.Anthropic, file_path: str) -> str:
    """Upload a file and return its file_id."""
    file_obj = client.beta.files.upload(
        file=Path(file_path),
        betas=["files-api-2025-04-14"]
    )
    return file_obj.id


def run_agent(
    client: anthropic.Anthropic,
    docx_path: str,
    data_path: str,
    syntax_path: str,
    output_path: str,
    model: str = "claude-sonnet-4-5-20250929",
    max_iterations: int = 15,
    cost_limit: float = 10.0
):
    """Run the agent with files in container."""
    
    # Upload all files first
    print("Uploading files...")
    docx_file_id = upload_file(client, docx_path)
    print(f"  Uploaded {docx_path} -> {docx_file_id}")
    
    data_file_id = upload_file(client, data_path)
    print(f"  Uploaded {data_path} -> {data_file_id}")
    
    syntax_file_id = upload_file(client, syntax_path)
    print(f"  Uploaded {syntax_path} -> {syntax_file_id}")
    
    # System prompt
    system_prompt = """You are an expert at converting Word documents into Ketryx templates.

## Reference Files in Container
The following files are available in your working directory:
- ketryx_data.json - Item types, fields, relations available in Ketryx
- ketryx_syntax.json - Template syntax reference

Read these files using code execution when you need to look up:
- Exact field names (they use underscores, specific capitalization)
- Available relation types (affects, implements, tests, etc.)
- Valid status values
- Item type names (Anomaly, Requirement, etc.)

## Quick Syntax Reference
{project.name}                    Project name
{version.name}                    Version name
{@toc}                           Table of contents
{$KQL var = type:X status:Y}     Query items
{#var}...{/var}                  Loop
{docId}, {title}                 Item fields
{fieldValue.Field_Name}          Custom fields (underscores for spaces)
{relations | where:'type == "X"' | map:'other.docId' | join:', '}

## Your Task
1. Read the reference files to understand available fields/syntax
2. Analyze the document structure
3. Replace dynamic content with template variables
4. Use KQL + loops for data tables
5. Save the template, preserving ALL formatting"""

    # First message references uploaded files via container_upload
    initial_content = [
        {"type": "container_upload", "file_id": docx_file_id},
        {"type": "container_upload", "file_id": data_file_id},
        {"type": "container_upload", "file_id": syntax_file_id},
        {
            "type": "text",
            "text": f"""Convert the attached Word document (input_document.docx) into a Ketryx template.

I've included two reference files:
- ketryx_data.json: Contains all available item types, fields, and relations
- ketryx_syntax.json: Contains the complete templating syntax reference

Please:
1. First, read the reference files to understand what fields and syntax are available
2. Analyze the document to identify dynamic content (project names, versions, defect tables, counts)
3. Replace dynamic content with appropriate template variables
4. For tables with repeated data rows, add KQL queries and loops
5. Save the result to: {output_path}

The output must preserve ALL formatting - only replace text content, not structure."""
        }
    ]
    
    messages = [{"role": "user", "content": initial_content}]
    
    print("\n" + "="*60)
    print("KETRYX TEMPLATE AGENT v6 (File-Based Context)")
    print("="*60)
    print(f"Model: {model}")
    print(f"Cost limit: ${cost_limit:.2f}")
    print(f"Document: {docx_path}")
    print("="*60)
    
    total_cost = 0.0
    container_id = None
    iteration = 0
    
    while iteration < max_iterations:
        iteration += 1
        print(f"\n--- Iteration {iteration} ---")
        
        if total_cost >= cost_limit:
            print(f"\n⚠️ Cost limit reached: ${total_cost:.2f}")
            return {"success": False, "error": "Cost limit reached", "total_cost": total_cost}
        
        container_config = {
            "skills": [{"type": "anthropic", "skill_id": "docx", "version": "latest"}]
        }
        if container_id:
            container_config["id"] = container_id
        
        print("Processing...", end="", flush=True)
        
        try:
            response = client.beta.messages.create(
                model=model,
                max_tokens=8000,
                betas=["code-execution-2025-08-25", "skills-2025-10-02", "files-api-2025-04-14"],
                system=system_prompt,
                container=container_config,
                messages=messages,
                tools=[{"type": "code_execution_20250825", "name": "code_execution"}]
            )
        except anthropic.APIError as e:
            print(f"\nAPI Error: {e}")
            return {"success": False, "error": str(e), "total_cost": total_cost}
        
        print(" done.")
        
        usage = response.usage
        iter_cost = estimate_cost(model, usage.input_tokens, usage.output_tokens)
        total_cost += iter_cost
        
        print(f"  Tokens: {usage.input_tokens:,} in / {usage.output_tokens:,} out")
        print(f"  Cost: ${iter_cost:.3f} (total: ${total_cost:.3f})")
        
        if hasattr(response, 'container') and response.container:
            container_id = response.container.id
        
        has_file = False
        for block in response.content:
            if block.type == "text" and block.text.strip():
                text = block.text[:400] + "..." if len(block.text) > 400 else block.text
                print(f"\nClaude: {text}")
            
            elif block.type == "bash_code_execution_tool_result":
                result_content = getattr(block, 'content', None)
                if result_content:
                    inner_content = getattr(result_content, 'content', [])
                    for item in inner_content:
                        if hasattr(item, 'file_id'):
                            has_file = True
                            file_id = item.file_id
                            print(f"\n> Generated file: {file_id}")
                            
                            try:
                                meta = client.beta.files.retrieve_metadata(
                                    file_id=file_id,
                                    betas=["files-api-2025-04-14"]
                                )
                                print(f"  Name: {meta.filename}")
                                
                                file_data = client.beta.files.download(
                                    file_id=file_id,
                                    betas=["files-api-2025-04-14"]
                                )
                                
                                out_path = Path(output_path)
                                file_data.write_to_file(str(out_path))
                                print(f"  Saved: {out_path} ({out_path.stat().st_size:,} bytes)")
                                
                            except Exception as e:
                                print(f"  Error: {e}")
        
        if response.stop_reason == "end_turn":
            if has_file:
                print(f"\n{'='*60}")
                print(f"✓ COMPLETE")
                print(f"  Output: {output_path}")
                print(f"  Iterations: {iteration}")
                print(f"  Total cost: ${total_cost:.3f}")
                print(f"{'='*60}")
                return {"success": True, "output_path": output_path, "total_cost": total_cost}
            else:
                print("\nNo file output, may need to continue...")
                messages.append({"role": "assistant", "content": response.content})
                messages.append({
                    "role": "user", 
                    "content": "Please save the template document now."
                })
                continue
        
        elif response.stop_reason == "pause_turn":
            print("  (Continuing long operation...)")
            messages.append({"role": "assistant", "content": response.content})
            continue
        
        else:
            messages.append({"role": "assistant", "content": response.content})
    
    print(f"\nStopped after {iteration} iterations")
    return {"success": False, "error": "Max iterations", "total_cost": total_cost}


def main():
    parser = argparse.ArgumentParser(description="Ketryx Template Generator v6")
    parser.add_argument("--docx", required=True)
    parser.add_argument("--data", required=True)
    parser.add_argument("--syntax", required=True)
    parser.add_argument("--output", required=True)
    parser.add_argument("--model", default="claude-sonnet-4-5-20250929")
    parser.add_argument("--max-iterations", type=int, default=15)
    parser.add_argument("--cost-limit", type=float, default=10.0)
    
    args = parser.parse_args()
    
    for p, n in [(args.docx, "docx"), (args.data, "data"), (args.syntax, "syntax")]:
        if not Path(p).exists():
            print(f"Error: {n} not found: {p}")
            sys.exit(1)
    
    client = anthropic.Anthropic()
    
    result = run_agent(
        client=client,
        docx_path=args.docx,
        data_path=args.data,
        syntax_path=args.syntax,
        output_path=args.output,
        model=args.model,
        max_iterations=args.max_iterations,
        cost_limit=args.cost_limit
    )
    
    print(f"\nFinal cost: ${result.get('total_cost', 0):.3f}")
    
    if not result.get("success"):
        print(f"Failed: {result.get('error')}")
        sys.exit(1)


if __name__ == "__main__":
    main()