#!/usr/bin/env python
"""
Documentation Management Utility

This script provides utilities for managing the AR Data Analysis documentation system.
It can generate indexes, validate documentation, and help maintain the documentation structure.
"""

import os
import json
import re
import datetime
import argparse
from typing import Dict, List, Any, Optional


class DocumentationManager:
    """Manager for AR Data Analysis documentation system."""

    def __init__(self, docs_dir: str = "."):
        """
        Initialize the documentation manager.
        
        Args:
            docs_dir: Root directory of documentation
        """
        self.docs_dir = docs_dir
        self.config_path = os.path.join(docs_dir, "docs_config.json")
        self.readme_path = os.path.join(docs_dir, "README.md")
        self.config = self._load_config()
        
    def _load_config(self) -> Dict:
        """Load the documentation configuration."""
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Error: Configuration file not found at {self.config_path}")
            return {
                "project_name": "AR Data Analysis",
                "categories": [],
                "documents": []
            }
    
    def save_config(self):
        """Save the current configuration."""
        with open(self.config_path, 'w') as f:
            json.dump(self.config, f, indent=2)
        print(f"Configuration saved to {self.config_path}")
    
    def scan_documents(self) -> List[Dict]:
        """
        Scan the documentation directory for markdown files.
        
        Returns:
            List of document metadata dictionaries
        """
        documents = []
        
        for category in self.config["categories"]:
            category_path = os.path.join(self.docs_dir, category["path"])
            if not os.path.exists(category_path):
                continue
                
            for filename in os.listdir(category_path):
                if filename.endswith(".md"):
                    file_path = os.path.join(category_path, filename)
                    rel_path = os.path.join(category["path"], filename)
                    
                    # Extract title from markdown file
                    title = self._extract_title(file_path)
                    
                    # Check if document already exists in config
                    existing_doc = next((doc for doc in self.config["documents"] 
                                         if doc["path"] == rel_path), None)
                    
                    if existing_doc:
                        # Update existing document
                        existing_doc["title"] = title
                        existing_doc["last_updated"] = datetime.date.today().isoformat()
                        documents.append(existing_doc)
                    else:
                        # Create new document entry
                        new_doc = {
                            "title": title,
                            "path": rel_path,
                            "category": category["name"],
                            "tags": [],
                            "created": datetime.date.today().isoformat(),
                            "last_updated": datetime.date.today().isoformat(),
                            "related_docs": []
                        }
                        documents.append(new_doc)
        
        return documents
    
    def _extract_title(self, file_path: str) -> str:
        """Extract the title from a markdown file."""
        try:
            with open(file_path, 'r') as f:
                content = f.read()
                # Look for the first heading
                match = re.search(r'^#\s+(.+)$', content, re.MULTILINE)
                if match:
                    return match.group(1)
                return os.path.basename(file_path)
        except Exception:
            return os.path.basename(file_path)
    
    def update_index(self):
        """Update the documentation index in README.md."""
        # Scan for documents
        documents = self.scan_documents()
        
        # Update config with scanned documents
        self.config["documents"] = documents
        self.save_config()
        
        # Generate README content
        readme_content = f"""# {self.config['project_name']} Documentation

Welcome to the {self.config['project_name']} documentation. This repository contains comprehensive documentation for the {self.config['project_name']} project, including technical guides, troubleshooting information, and architectural overviews.

## Documentation Structure

The documentation is organized into the following categories:

"""
        
        # Add categories
        for category in self.config["categories"]:
            readme_content += f"- **[{category['name']}](./{category['path']}/)**: {category['description']}\n"
        
        # Add document index
        readme_content += """
## Documentation Index

| Document | Category | Description |
|----------|----------|-------------|
"""
        
        # Add document entries
        for doc in sorted(documents, key=lambda x: x["category"]):
            # Extract description from file
            description = self._extract_description(os.path.join(self.docs_dir, doc["path"]))
            readme_content += f"| [{doc['title']}](./{doc['path']}) | {doc['category']} | {description} |\n"
        
        # Add usage information
        readme_content += """
## Using This Documentation

This documentation is designed to be:

1. **Modular**: Each document focuses on a specific topic or issue
2. **Indexed**: All documents are cataloged in this central index
3. **Cross-referenced**: Documents link to related content where appropriate
4. **Searchable**: Use the search function to find specific information

## Contributing to Documentation

When adding new documentation:

1. Place it in the appropriate category folder
2. Add an entry to the index table in this README
3. Follow the established Markdown formatting guidelines
4. Include links to related documents where relevant

## Documentation Standards

All documentation should follow these standards:

- Use Markdown formatting for consistency
- Include a clear title and description
- Organize content with appropriate headings
- Include code examples where relevant
- Link to related documentation
- Include version information where appropriate
"""
        
        # Write README
        with open(self.readme_path, 'w') as f:
            f.write(readme_content)
        
        print(f"Documentation index updated in {self.readme_path}")
    
    def _extract_description(self, file_path: str) -> str:
        """Extract a description from a markdown file."""
        try:
            with open(file_path, 'r') as f:
                content = f.read()
                # Skip the title and look for the first paragraph
                content_without_title = re.sub(r'^#.*$', '', content, 1, re.MULTILINE)
                match = re.search(r'\n\n(.+?)\n\n', content_without_title, re.DOTALL)
                if match:
                    # Truncate and clean up description
                    desc = match.group(1).replace('\n', ' ').strip()
                    if len(desc) > 100:
                        desc = desc[:97] + '...'
                    return desc
                return "No description available"
        except Exception:
            return "No description available"
    
    def create_document_template(self, title: str, category: str) -> Optional[str]:
        """
        Create a new document from template.
        
        Args:
            title: Document title
            category: Document category
            
        Returns:
            Path to the created document or None if failed
        """
        # Validate category
        category_obj = next((cat for cat in self.config["categories"] 
                            if cat["name"].lower() == category.lower()), None)
        if not category_obj:
            print(f"Error: Category '{category}' not found")
            return None
        
        # Create filename
        filename = title.lower().replace(' ', '_') + '.md'
        category_path = os.path.join(self.docs_dir, category_obj["path"])
        file_path = os.path.join(category_path, filename)
        
        # Ensure directory exists
        os.makedirs(category_path, exist_ok=True)
        
        # Create document from template
        today = datetime.date.today().isoformat()
        content = f"""# {title}

Brief description of the document.

## Overview

Detailed overview of the topic.

## Details

Main content of the document.

## Related Documentation

- [Link to related document](../path/to/document.md)

## Version Information

- Created: {today}
- Last Updated: {today}
- Applies to: AR Data Analysis v1.0.0+
"""
        
        with open(file_path, 'w') as f:
            f.write(content)
        
        print(f"Created document template at {file_path}")
        
        # Update index
        self.update_index()
        
        return os.path.join(category_obj["path"], filename)
    
    def validate_documentation(self) -> bool:
        """
        Validate the documentation structure and links.
        
        Returns:
            True if validation passed, False otherwise
        """
        all_valid = True
        
        # Check for missing documents
        for doc in self.config["documents"]:
            doc_path = os.path.join(self.docs_dir, doc["path"])
            if not os.path.exists(doc_path):
                print(f"Error: Document {doc['path']} referenced in config but not found")
                all_valid = False
        
        # Check for broken internal links
        for doc in self.config["documents"]:
            doc_path = os.path.join(self.docs_dir, doc["path"])
            if os.path.exists(doc_path):
                with open(doc_path, 'r') as f:
                    content = f.read()
                    # Find all markdown links
                    links = re.findall(r'\[.+?\]\((.+?)\)', content)
                    for link in links:
                        # Skip external links
                        if link.startswith(('http://', 'https://')):
                            continue
                        
                        # Check if internal link exists
                        link_path = os.path.normpath(os.path.join(
                            os.path.dirname(doc_path), link))
                        if not os.path.exists(link_path):
                            print(f"Warning: Broken link in {doc['path']}: {link}")
                            all_valid = False
        
        if all_valid:
            print("Documentation validation passed")
        
        return all_valid


def main():
    """Main entry point for the documentation manager."""
    parser = argparse.ArgumentParser(description='AR Data Analysis Documentation Manager')
    parser.add_argument('--docs-dir', default='.', help='Documentation root directory')
    
    subparsers = parser.add_subparsers(dest='command', help='Command to run')
    
    # Update index command
    subparsers.add_parser('update-index', help='Update the documentation index')
    
    # Create document command
    create_parser = subparsers.add_parser('create', help='Create a new document')
    create_parser.add_argument('title', help='Document title')
    create_parser.add_argument('category', help='Document category')
    
    # Validate command
    subparsers.add_parser('validate', help='Validate documentation')
    
    args = parser.parse_args()
    
    manager = DocumentationManager(args.docs_dir)
    
    if args.command == 'update-index':
        manager.update_index()
    elif args.command == 'create':
        manager.create_document_template(args.title, args.category)
    elif args.command == 'validate':
        manager.validate_documentation()
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
