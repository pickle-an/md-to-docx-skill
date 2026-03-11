#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
import os
from typing import Tuple, Optional


class VersionManager:
    VERSION_PATTERN = re.compile(r'(.+?)[-_]?V(\d+)$', re.IGNORECASE)
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
    
    def extract_version(self, filename: str) -> Tuple[str, Optional[int]]:
        base_name = os.path.splitext(filename)[0]
        match = self.VERSION_PATTERN.match(base_name)
        
        if match:
            name_part = match.group(1).rstrip('_-')
            version_num = int(match.group(2))
            return name_part, version_num
        
        return base_name, None
    
    def get_next_version(self, filename: str) -> int:
        _, current_version = self.extract_version(filename)
        
        if current_version is None:
            return 1
        
        return current_version + 1
    
    def generate_versioned_filename(
        self,
        base_path: str,
        extension: str,
        version: Optional[int] = None,
        suffix: str = ''
    ) -> str:
        dir_path = os.path.dirname(base_path)
        filename = os.path.basename(base_path)
        name_without_ext = os.path.splitext(filename)[0]
        
        name_part, existing_version = self.extract_version(filename)
        
        if version is not None:
            final_version = version
        elif existing_version is not None:
            final_version = existing_version
        else:
            final_version = 1
        
        if suffix:
            new_name = f"{name_part}_V{final_version}_{suffix}.{extension}"
        else:
            new_name = f"{name_part}_V{final_version}.{extension}"
        
        if dir_path:
            return os.path.join(dir_path, new_name)
        return new_name
    
    def generate_next_versioned_filename(
        self,
        base_path: str,
        extension: str,
        suffix: str = ''
    ) -> Tuple[str, int]:
        dir_path = os.path.dirname(base_path)
        filename = os.path.basename(base_path)
        name_without_ext = os.path.splitext(filename)[0]
        
        name_part, existing_version = self.extract_version(filename)
        
        if existing_version is not None:
            next_version = existing_version + 1
        else:
            next_version = 1
        
        if suffix:
            new_name = f"{name_part}_V{next_version}_{suffix}.{extension}"
        else:
            new_name = f"{name_part}_V{next_version}.{extension}"
        
        if dir_path:
            return os.path.join(dir_path, new_name), next_version
        return new_name, next_version
    
    def find_latest_version_file(
        self,
        base_path: str,
        extension: str
    ) -> Tuple[Optional[str], int]:
        dir_path = os.path.dirname(base_path)
        filename = os.path.basename(base_path)
        name_without_ext = os.path.splitext(filename)[0]
        
        name_part, existing_version = self.extract_version(filename)
        
        if dir_path:
            search_dir = dir_path
        else:
            search_dir = '.'
        
        if not os.path.exists(search_dir):
            return None, 0
        
        max_version = 0
        latest_file = None
        
        pattern = re.compile(
            re.escape(name_part) + r'[-_]?V(\d+)' + re.escape('.' + extension) + '$',
            re.IGNORECASE
        )
        
        for f in os.listdir(search_dir):
            match = pattern.match(f)
            if match:
                version_num = int(match.group(1))
                if version_num > max_version:
                    max_version = version_num
                    latest_file = os.path.join(search_dir, f)
        
        if existing_version is not None and existing_version > max_version:
            return base_path, existing_version
        
        return latest_file, max_version
    
    def get_next_version_from_directory(
        self,
        base_path: str,
        extension: str
    ) -> Tuple[str, int]:
        latest_file, max_version = self.find_latest_version_file(base_path, extension)
        
        next_version = max_version + 1
        
        dir_path = os.path.dirname(base_path)
        filename = os.path.basename(base_path)
        name_without_ext = os.path.splitext(filename)[0]
        
        name_part, _ = self.extract_version(filename)
        
        new_name = f"{name_part}_V{next_version}.{extension}"
        
        if dir_path:
            return os.path.join(dir_path, new_name), next_version
        return new_name, next_version


def get_versioned_output_paths(
    input_path: str,
    output_dir: Optional[str] = None,
    verbose: bool = False
) -> dict:
    manager = VersionManager(verbose=verbose)
    
    input_dir = os.path.dirname(input_path)
    input_filename = os.path.basename(input_path)
    name_without_ext = os.path.splitext(input_filename)[0]
    
    name_part, existing_version = manager.extract_version(input_filename)
    
    if output_dir is None:
        output_dir = input_dir if input_dir else '.'
    
    md_pattern = re.compile(
        re.escape(name_part) + r'[-_]?V(\d+)_normalized\.md$',
        re.IGNORECASE
    )
    docx_pattern = re.compile(
        re.escape(name_part) + r'[-_]?V(\d+)\.docx$',
        re.IGNORECASE
    )
    
    max_md_version = 0
    max_docx_version = 0
    
    if os.path.exists(output_dir):
        for f in os.listdir(output_dir):
            md_match = md_pattern.match(f)
            if md_match:
                version_num = int(md_match.group(1))
                max_md_version = max(max_md_version, version_num)
            
            docx_match = docx_pattern.match(f)
            if docx_match:
                version_num = int(docx_match.group(1))
                max_docx_version = max(max_docx_version, version_num)
    
    if existing_version is not None:
        max_md_version = max(max_md_version, existing_version)
        max_docx_version = max(max_docx_version, existing_version)
    
    next_version = max(max_md_version, max_docx_version) + 1
    
    if existing_version is None and max_md_version == 0 and max_docx_version == 0:
        next_version = 1
    
    normalized_md_path = os.path.join(output_dir, f"{name_part}_V{next_version}_normalized.md")
    docx_path = os.path.join(output_dir, f"{name_part}_V{next_version}.docx")
    
    if verbose:
        print(f"[VersionManager] Input file: {input_path}")
        print(f"[VersionManager] Name part: {name_part}")
        print(f"[VersionManager] Existing version in filename: {existing_version}")
        print(f"[VersionManager] Max MD version in directory: {max_md_version}")
        print(f"[VersionManager] Max DOCX version in directory: {max_docx_version}")
        print(f"[VersionManager] Next version: {next_version}")
        print(f"[VersionManager] Normalized MD path: {normalized_md_path}")
        print(f"[VersionManager] DOCX path: {docx_path}")
    
    return {
        'name_part': name_part,
        'version': next_version,
        'normalized_md_path': normalized_md_path,
        'docx_path': docx_path
    }


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python version_manager.py <input_file> [output_dir]")
        print("\nExamples:")
        print("  python version_manager.py document.md")
        print("  python version_manager.py document_V2.md")
        print("  python version_manager.py document.md ./output")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    result = get_versioned_output_paths(input_file, output_dir, verbose=True)
    
    print(f"\nResult:")
    print(f"  Version: V{result['version']}")
    print(f"  Normalized MD: {result['normalized_md_path']}")
    print(f"  DOCX: {result['docx_path']}")
