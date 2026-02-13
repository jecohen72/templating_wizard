#!/usr/bin/env python3
"""
Ketryx Project Data Extractor v2

Extracts project data from Ketryx API and builds an AI-optimized data model
for template generation. Pre-computes all deterministic access paths.

Usage:
    python ketryx_data_extractor.py --project-id KXPRJ... --api-key YOUR_KEY
    
Environment variables:
    KETRYX_API_KEY: API key (alternative to --api-key)
    KETRYX_BASE_URL: Base URL (default: https://app.ketryx.com)
"""

import argparse
import json
import os
import sys
import time
from collections import defaultdict
from dataclasses import dataclass, field
from typing import Any, Optional, Union, List, Dict
from urllib.parse import urljoin

import requests


# =============================================================================
# Configuration
# =============================================================================

DEFAULT_BASE_URL = "https://app.ketryx.com"
MAX_SAMPLE_VALUES = 5
MAX_UNIQUE_VALUES = 25
MAX_SAMPLE_RECORDS = 3
MAX_STRING_LENGTH = 100
REQUEST_DELAY_SECONDS = 0.1

# Fields that commonly contain rich text
RICH_TEXT_FIELD_PATTERNS = {
    "description", "rationale", "notes", "content", "details", "summary",
    "comment", "justification", "analysis", "resolution", "root_cause",
    "impact", "mitigation", "verification", "acceptance_criteria",
}

# Built-in template variables
BUILTIN_VARIABLES = {
    "project": {
        "project.name": {"description": "Project name", "dataType": "string", "access": "project.name"},
        "project.id": {"description": "Project Ketryx ID", "dataType": "string", "access": "project.id"},
    },
    "version": {
        "version.name": {"description": "Version name/number", "dataType": "string", "access": "version.name"},
        "version.id": {"description": "Version Ketryx ID", "dataType": "string", "access": "version.id"},
        "version.isReleased": {"description": "Whether version is released", "dataType": "boolean", "access": "version.isReleased"},
        "version.releaseDate": {"description": "Release date (if released)", "dataType": "datetime", "access": "version.releaseDate"},
    },
    "document": {
        "document.title": {"description": "Document title", "dataType": "string", "access": "document.title"},
        "document.date": {"description": "Document generation date", "dataType": "datetime", "access": "document.date"},
        "document.version": {"description": "Document version", "dataType": "string", "access": "document.version"},
    },
}


# =============================================================================
# Data classes
# =============================================================================

@dataclass
class FieldInfo:
    name: str
    normalizedName: str
    label: str
    dataType: str
    isCustomField: bool
    access: dict
    description: str = ""
    uniqueValues: list = field(default_factory=list)
    exampleValues: list = field(default_factory=list)
    fillRate: float = 0.0
    
    def to_dict(self) -> dict:
        result = {
            "label": self.label,
            "dataType": self.dataType,
            "access": self.access,
        }
        if self.isCustomField:
            result["isCustomField"] = True
        if self.uniqueValues:
            result["uniqueValues"] = self.uniqueValues
        elif self.exampleValues:
            result["exampleValues"] = self.exampleValues
        if self.fillRate < 100:
            result["fillRate"] = f"{self.fillRate:.0f}%"
        return result


@dataclass
class RelationTypeInfo:
    relationType: str
    fromTypes: list = field(default_factory=list)
    toTypes: list = field(default_factory=list)
    count: int = 0
    accessPattern: str = ""
    
    def to_dict(self) -> dict:
        return {
            "relationType": self.relationType,
            "fromTypes": list(set(self.fromTypes)),
            "toTypes": list(set(self.toTypes)),
            "count": self.count,
            "accessPattern": self.accessPattern,
        }


@dataclass
class ItemTypeInfo:
    name: str
    kqlQuery: str
    shortName: str = ""
    count: int = 0
    statuses: list = field(default_factory=list)
    fields: dict = field(default_factory=dict)
    sampleRecords: list = field(default_factory=list)
    outgoingRelations: list = field(default_factory=list)
    incomingRelations: list = field(default_factory=list)
    
    def to_dict(self) -> dict:
        result = {
            "kqlQuery": self.kqlQuery,
            "count": self.count,
            "fields": {name: info.to_dict() for name, info in self.fields.items()},
            "sampleRecords": self.sampleRecords,
        }
        if self.shortName:
            result["shortName"] = self.shortName
        if self.statuses:
            result["statuses"] = self.statuses
        if self.outgoingRelations:
            result["outgoingRelations"] = self.outgoingRelations
        if self.incomingRelations:
            result["incomingRelations"] = self.incomingRelations
        return result


# =============================================================================
# API Client
# =============================================================================

class KetryxAPIClient:
    """Client for Ketryx API."""
    
    def __init__(self, base_url: str, api_key: str):
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Bearer {api_key}",
            "Accept": "application/json",
        })
    
    def _request(self, method: str, endpoint: str, **kwargs) -> Optional[Union[dict, list]]:
        url = urljoin(self.base_url + "/", endpoint.lstrip("/"))
        try:
            response = self.session.request(method, url, **kwargs)
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            print(f"HTTP Error {e.response.status_code} for {endpoint}", file=sys.stderr)
            if e.response.status_code in [400, 403, 404]:
                print(f"  Response: {e.response.text[:300]}", file=sys.stderr)
            return None
        except requests.exceptions.RequestException as e:
            print(f"Request failed for {endpoint}: {e}", file=sys.stderr)
            return None
    
    def get_project(self, project_id: str) -> Optional[dict]:
        return self._request("GET", f"/api/v1/projects/{project_id}")
    
    def get_versions(self, project_id: str) -> Optional[dict]:
        return self._request("GET", f"/api/v1/projects/{project_id}/versions")
    
    def get_items(self, project_id: str, start_at: int = 0, max_results: int = 1000) -> Optional[dict]:
        return self._request("GET", f"/api/v1/projects/{project_id}/items", 
                           params={"startAt": start_at, "maxResults": max_results})
    
    def get_item_records(self, project_id: str, item_id: str) -> Optional[dict]:
        """Get records for a specific item."""
        return self._request("GET", f"/api/v1/projects/{project_id}/items/{item_id}/records")
    
    def query_records(self, project_id: str, kql: str, version_id: str = None, 
                     start_at: int = 0, max_results: int = 1000) -> Optional[dict]:
        """Query records using KQL."""
        params = {"query": kql, "startAt": start_at, "maxResults": max_results}
        if version_id:
            params["versionId"] = version_id
        return self._request("GET", f"/api/v1/projects/{project_id}/records", params=params)


# =============================================================================
# Data Extractor
# =============================================================================

class KetryxDataExtractor:
    """Extracts and processes project data into AI-optimized format."""
    
    def __init__(self, client: KetryxAPIClient, project_id: str, version_id: str = None):
        self.client = client
        self.project_id = project_id
        self.version_id = version_id
        self.item_types: Dict[str, ItemTypeInfo] = {}
        self.relation_types: Dict[str, RelationTypeInfo] = {}
        self.all_records_by_type: Dict[str, list] = {}
    
    def extract(self) -> dict:
        """Main extraction workflow."""
        
        # 1. Get project metadata
        print("Fetching project metadata...", file=sys.stderr)
        project = self.client.get_project(self.project_id)
        if not project:
            raise RuntimeError("Failed to fetch project")
        
        # 2. Get versions
        print("Fetching versions...", file=sys.stderr)
        versions_response = self.client.get_versions(self.project_id)
        versions = versions_response.get("versions", []) if isinstance(versions_response, dict) else []
        
        # Auto-select version if not specified
        if not self.version_id and versions:
            released = [v for v in versions if v.get("isReleased")]
            self.version_id = released[-1]["id"] if released else versions[-1]["id"]
        
        current_version = next((v for v in versions if v.get("id") == self.version_id), None)
        
        # 3. Discover item types by sampling items
        print("Discovering item types...", file=sys.stderr)
        self._discover_types_from_items()
        
        # 4. Query records for each discovered type
        print("Fetching records by type...", file=sys.stderr)
        self._fetch_records_by_type()
        
        # 5. Analyze fields for each type
        print("Analyzing fields...", file=sys.stderr)
        self._analyze_all_fields()
        
        # 6. Analyze relations
        print("Analyzing relations...", file=sys.stderr)
        self._analyze_relations()
        
        # 7. Build output
        return self._build_output(project, versions, current_version)
    
    def _discover_types_from_items(self):
        """Discover item types by fetching records for sample items."""
        # Get items
        items_response = self.client.get_items(self.project_id, max_results=100)
        items = items_response.get("items", []) if isinstance(items_response, dict) else []
        
        print(f"  Sampling {min(len(items), 50)} items to discover types...", file=sys.stderr)
        
        discovered_types = set()
        
        # Sample items to discover types
        for item in items[:50]:
            if not isinstance(item, dict):
                continue
            
            item_id = item.get("id")
            if not item_id:
                continue
            
            time.sleep(REQUEST_DELAY_SECONDS)
            
            records_response = self.client.get_item_records(self.project_id, item_id)
            if not records_response:
                continue
            
            records = records_response.get("records", [])
            for record in records:
                if isinstance(record, dict):
                    type_name = record.get("type", "")
                    if type_name:
                        discovered_types.add(type_name)
        
        print(f"  Discovered types: {discovered_types}", file=sys.stderr)
        
        # Create ItemTypeInfo for each discovered type
        for type_name in discovered_types:
            if not type_name or type_name == "Unknown":
                continue
            
            # Build KQL query (quote if contains spaces)
            kql = f'type:"{type_name}"' if " " in type_name else f"type:{type_name}"
            
            self.item_types[type_name] = ItemTypeInfo(
                name=type_name,
                kqlQuery=kql,
                shortName="",  # Will infer from docId later
            )
    
    def _fetch_records_by_type(self):
        """Fetch all records for each discovered type."""
        types_to_remove = []
        
        for type_name, type_info in self.item_types.items():
            time.sleep(REQUEST_DELAY_SECONDS)
            
            # Don't pass versionId - query all records for this type
            records_response = self.client.query_records(
                self.project_id,
                type_info.kqlQuery,
                version_id=None,  # Skip version filtering
                max_results=1000
            )
            
            if not records_response:
                types_to_remove.append(type_name)
                continue
            
            records = records_response.get("records", [])
            total = records_response.get("total", len(records))
            
            if records:
                self.all_records_by_type[type_name] = records
                type_info.count = total
                
                # Infer short name from docId patterns
                for record in records[:20]:
                    if isinstance(record, dict):
                        # Look for docId in fields
                        for field_obj in record.get("fields", []):
                            if isinstance(field_obj, dict) and field_obj.get("label") == "ID":
                                doc_id = field_obj.get("value", "")
                                if doc_id and "-" in str(doc_id):
                                    prefix = str(doc_id).split("-")[0]
                                    if prefix.isalpha() and 1 <= len(prefix) <= 6:
                                        type_info.shortName = prefix
                                        break
                        if type_info.shortName:
                            break
                
                print(f"  {type_name}: {len(records)} records (total: {total})", file=sys.stderr)
            else:
                types_to_remove.append(type_name)
        
        # Remove types with no records
        for type_name in types_to_remove:
            del self.item_types[type_name]
    
    def _analyze_all_fields(self):
        """Analyze fields for all item types."""
        for type_name, records in self.all_records_by_type.items():
            if not records:
                continue
            
            type_info = self.item_types.get(type_name)
            if not type_info:
                continue
            
            # Collect field data
            field_data = defaultdict(lambda: {"values": [], "types": set(), "isCustom": True, "label": ""})
            
            # Standard fields from record root
            standard_fields = ["title", "revision", "isControlled", "createdAt"]
            
            for record in records:
                if not isinstance(record, dict):
                    continue
                
                # Standard fields
                for field_name in standard_fields:
                    value = record.get(field_name)
                    if value is not None and value != "":
                        field_data[field_name]["values"].append(value)
                        field_data[field_name]["isCustom"] = False
                        field_data[field_name]["label"] = field_name
                
                # Custom fields from fields array
                for field_obj in record.get("fields", []):
                    if not isinstance(field_obj, dict):
                        continue
                    
                    label = field_obj.get("label", "")
                    value = field_obj.get("value")
                    field_type = field_obj.get("type", "string")
                    
                    if label and value is not None and value != "":
                        normalized = self._normalize_field_name(label)
                        field_data[normalized]["values"].append(value)
                        field_data[normalized]["types"].add(field_type)
                        field_data[normalized]["isCustom"] = True
                        field_data[normalized]["label"] = label
            
            # Build FieldInfo objects
            for field_key, data in field_data.items():
                is_custom = data.get("isCustom", True)
                label = data.get("label", field_key)
                values = data["values"]
                
                if not values:
                    continue
                
                # Determine data type
                types = data["types"]
                if self._is_rich_text_field(field_key, values):
                    data_type = "richText"
                elif "number" in types:
                    data_type = "number"
                elif "date" in types or "datetime" in types:
                    data_type = "datetime"
                elif "boolean" in types:
                    data_type = "boolean"
                else:
                    data_type = "string"
                
                # Compute access paths
                access = self._compute_access_paths(field_key, is_custom, data_type)
                
                # Compute unique/example values
                unique_values, example_values = self._compute_value_samples(values, data_type)
                
                # Compute fill rate
                fill_rate = (len(values) / len(records) * 100) if records else 0
                
                type_info.fields[field_key] = FieldInfo(
                    name=field_key,
                    normalizedName=self._normalize_field_name(field_key),
                    label=label,
                    dataType=data_type,
                    isCustomField=is_custom,
                    access=access,
                    uniqueValues=unique_values,
                    exampleValues=example_values,
                    fillRate=fill_rate,
                )
            
            # Extract statuses from fields
            for record in records:
                for field_obj in record.get("fields", []):
                    if isinstance(field_obj, dict) and field_obj.get("label") == "Status":
                        status = field_obj.get("value")
                        if status and status not in type_info.statuses:
                            type_info.statuses.append(status)
            
            type_info.statuses = sorted(type_info.statuses)
            
            # Extract sample records
            type_info.sampleRecords = self._build_sample_records(records, type_info.fields)
    
    def _normalize_field_name(self, name: str) -> str:
        """Normalize field name: spaces -> underscores."""
        return name.replace(" ", "_")
    
    def _is_rich_text_field(self, field_name: str, values: list) -> bool:
        """Determine if a field contains rich text."""
        lower_name = field_name.lower().replace("_", " ")
        
        for pattern in RICH_TEXT_FIELD_PATTERNS:
            if pattern in lower_name:
                return True
        
        html_indicators = ["<p>", "<ul>", "<ol>", "<li>", "<strong>", "<em>", "<br", "<div>", "<span>"]
        for value in values[:10]:
            if isinstance(value, str):
                if any(tag in value for tag in html_indicators):
                    return True
                if len(value) > 500:
                    return True
        
        return False
    
    def _compute_access_paths(self, field_name: str, is_custom: bool, data_type: str) -> dict:
        """Compute all valid access paths for a field."""
        normalized = self._normalize_field_name(field_name)
        
        if is_custom:
            if data_type == "richText":
                return {
                    "plain": f"fieldValue.{normalized}",
                    "rich": f"~~fieldContent.{normalized}",
                }
            else:
                return {
                    "plain": f"fieldValue.{normalized}",
                }
        else:
            if data_type == "richText":
                return {
                    "plain": field_name,
                    "rich": f"~~{field_name}",
                }
            else:
                return {
                    "plain": field_name,
                }
    
    def _compute_value_samples(self, values: list, data_type: str) -> tuple:
        """Compute unique values (for categorical) or example values (for others)."""
        if not values:
            return [], []
        
        if data_type == "richText":
            return [], []
        
        truncated = []
        for v in values:
            if isinstance(v, str):
                truncated.append(v[:MAX_STRING_LENGTH])
            else:
                truncated.append(str(v))
        
        unique = list(dict.fromkeys(truncated))
        
        if len(unique) <= MAX_UNIQUE_VALUES:
            return sorted(unique), []
        else:
            return [], unique[:MAX_SAMPLE_VALUES]
    
    def _build_sample_records(self, records: list, fields: dict) -> list:
        """Build sample records for AI pattern matching."""
        samples = []
        
        for record in records[:MAX_SAMPLE_RECORDS]:
            if not isinstance(record, dict):
                continue
            
            sample = {}
            
            # Standard fields
            if record.get("title"):
                sample["title"] = self._truncate(str(record["title"]), 80)
            
            # Custom fields
            custom_count = 0
            for field_obj in record.get("fields", []):
                if custom_count >= 4:
                    break
                if not isinstance(field_obj, dict):
                    continue
                
                label = field_obj.get("label", "")
                value = field_obj.get("value")
                normalized = self._normalize_field_name(label)
                
                if normalized in fields and fields[normalized].dataType == "richText":
                    continue
                
                if label and value and value != "":
                    sample[normalized] = self._truncate(str(value), 60)
                    custom_count += 1
            
            if sample:
                samples.append(sample)
        
        return samples
    
    def _truncate(self, text: str, max_len: int) -> str:
        if len(text) <= max_len:
            return text
        return text[:max_len - 3] + "..."
    
    def _analyze_relations(self):
        """Analyze relations between items."""
        relation_data = defaultdict(lambda: {"fromTypes": [], "toTypes": [], "count": 0})
        
        for type_name, records in self.all_records_by_type.items():
            for record in records:
                if not isinstance(record, dict):
                    continue
                
                for relation in record.get("relations", []):
                    if not isinstance(relation, dict):
                        continue
                    
                    rel_type = relation.get("type", "UNKNOWN")
                    to_item = relation.get("toItem", {})
                    
                    relation_data[rel_type]["fromTypes"].append(type_name)
                    relation_data[rel_type]["count"] += 1
        
        for rel_type, data in relation_data.items():
            access_pattern = f"relations | where('type', '{rel_type}')"
            
            self.relation_types[rel_type] = RelationTypeInfo(
                relationType=rel_type,
                fromTypes=data["fromTypes"],
                toTypes=data["toTypes"],
                count=data["count"],
                accessPattern=access_pattern,
            )
        
        # Populate per-type relation summaries
        for type_name, type_info in self.item_types.items():
            outgoing = []
            
            for rel_type, rel_info in self.relation_types.items():
                from_types_unique = list(set(rel_info.fromTypes))
                
                if type_name in from_types_unique:
                    outgoing.append({
                        "relation": rel_type,
                        "accessPattern": f"relations | where('type', '{rel_type}')",
                    })
            
            type_info.outgoingRelations = outgoing
    
    def _build_output(self, project: dict, versions: list, current_version: dict) -> dict:
        """Build the final output structure."""
        
        sorted_types = sorted(
            self.item_types.items(),
            key=lambda x: x[1].count,
            reverse=True
        )
        
        sorted_relations = sorted(
            self.relation_types.values(),
            key=lambda x: x.count,
            reverse=True
        )
        
        return {
            "_meta": {
                "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                "purpose": "AI-optimized project data for template generation",
                "version": "2.0.0",
                "notes": [
                    "access.plain: Use in table cells, inline text, anywhere plain string is needed",
                    "access.rich: Use when HTML rendering is desired (prefixed with ~~)",
                    "kqlQuery: Pre-computed query to fetch all items of this type",
                    "uniqueValues: All possible values (for categorical fields with <=25 values)",
                    "exampleValues: Sample values (for fields with >25 unique values)",
                ],
            },
            
            "project": {
                "id": project.get("id"),
                "name": project.get("name"),
                "description": self._truncate(project.get("description", ""), 200),
            },
            
            "version": {
                "id": current_version.get("id") if current_version else None,
                "name": current_version.get("name") if current_version else None,
                "isReleased": current_version.get("isReleased") if current_version else None,
            },
            
            "allVersions": [
                {
                    "id": v.get("id"),
                    "name": v.get("name"),
                    "isReleased": v.get("isReleased", False),
                }
                for v in versions
            ],
            
            "summary": {
                "totalItems": sum(t.count for t in self.item_types.values()),
                "itemTypeCount": len(self.item_types),
                "totalRelations": sum(r.count for r in self.relation_types.values()),
            },
            
            "builtinVariables": BUILTIN_VARIABLES,
            
            "itemTypes": {
                name: info.to_dict()
                for name, info in sorted_types
            },
            
            "relationTypes": [
                r.to_dict() for r in sorted_relations[:20]
            ],
        }


# =============================================================================
# Main
# =============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Extract Ketryx project data into AI-optimized format"
    )
    parser.add_argument("--project-id", required=True, help="Ketryx project ID (KXPRJ...)")
    parser.add_argument("--api-key", default=os.environ.get("KETRYX_API_KEY"), help="Ketryx API key")
    parser.add_argument("--base-url", default=os.environ.get("KETRYX_BASE_URL", DEFAULT_BASE_URL), help="Ketryx base URL")
    parser.add_argument("--version-id", help="Specific version ID (default: latest)")
    parser.add_argument("--output", "-o", default="project_data.json", help="Output file")
    
    args = parser.parse_args()
    
    if not args.api_key:
        print("Error: API key required. Use --api-key or set KETRYX_API_KEY", file=sys.stderr)
        sys.exit(1)
    
    client = KetryxAPIClient(args.base_url, args.api_key)
    extractor = KetryxDataExtractor(client, args.project_id, args.version_id)
    
    try:
        data = extractor.extract()
    except Exception as e:
        print(f"Extraction failed: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
    
    output_json = json.dumps(data, indent=2, ensure_ascii=False)
    
    if args.output == "-":
        print(output_json)
    else:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(output_json)
        print(f"\nOutput written to {args.output}", file=sys.stderr)
        print(f"Summary: {data['summary']['totalItems']} items across {data['summary']['itemTypeCount']} types", file=sys.stderr)


if __name__ == "__main__":
    main()