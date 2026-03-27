#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re
from typing import List, Tuple, Optional


class MarkdownNormalizer:
    
    def __init__(self, verbose: bool = False):
        self.verbose = verbose
        self.fixes_applied = []
        self.chinese_num_map = {
            '零': '0', '〇': '0',
            '一': '1', '二': '2', '三': '3', '四': '4',
            '五': '5', '六': '6', '七': '7', '八': '8',
            '九': '9', '十': '10', '十一': '11', '十二': '12',
            '十三': '13', '十四': '14', '十五': '15', '十六': '16',
            '十七': '17', '十八': '18', '十九': '19', '二十': '20'
        }
    
    def normalize(self, content: str) -> str:
        self.fixes_applied = []
        
        lines = content.split('\n')
        lines = self._fix_code_blocks(lines)
        lines = self._fix_tables(lines)
        lines = self._adjust_heading_levels(lines)
        lines = self._fix_headings(lines)
        lines = self._fix_lists(lines)
        lines = self._fix_horizontal_rules(lines)
        lines = self._fix_inline_formatting(lines)
        lines = self._fix_empty_lines(lines)
        lines = self._fix_paragraph_spacing(lines)
        lines = self._fix_leading_whitespace(lines)
        
        result = '\n'.join(lines)
        result = self._fix_trailing_whitespace(result)
        
        if self.verbose and self.fixes_applied:
            print(f"[Normalizer] Applied {len(self.fixes_applied)} fixes:")
            for fix in self.fixes_applied[:10]:
                print(f"  - {fix}")
            if len(self.fixes_applied) > 10:
                print(f"  ... and {len(self.fixes_applied) - 10} more")
        
        return result
    
    def _log_fix(self, message: str):
        self.fixes_applied.append(message)
    
    def _convert_chinese_number_in_heading(self, text: str) -> str:
        patterns = [
            (r'^([一二三四五六七八九十]+)[、．.]\s*', lambda m: self._replace_chinese_num(m)),
            (r'^\(([一二三四五六七八九十]+)\)\s*', lambda m: ''),
        ]
        
        for pattern, replacement in patterns:
            match = re.match(pattern, text)
            if match:
                chinese_num = match.group(1)
                if chinese_num in self.chinese_num_map:
                    arabic_num = self.chinese_num_map[chinese_num]
                    if pattern == patterns[0][0]:
                        new_text = re.sub(pattern, f'{arabic_num}. ', text)
                    else:
                        new_text = re.sub(pattern, '', text)
                    self._log_fix(f"Removed Chinese number in heading: '({chinese_num})'")
                    return new_text
        
        return text
    
    def _replace_chinese_num(self, match) -> str:
        chinese_num = match.group(1)
        return self.chinese_num_map.get(chinese_num, chinese_num)
    
    def _fix_code_blocks(self, lines: List[str]) -> List[str]:
        result = []
        in_code_block = False
        code_block_start = -1
        
        for i, line in enumerate(lines):
            if line.strip().startswith('```'):
                if not in_code_block:
                    in_code_block = True
                    code_block_start = i
                    result.append(line)
                else:
                    in_code_block = False
                    result.append(line)
            else:
                result.append(line)
        
        if in_code_block:
            result.append('```')
            self._log_fix(f"Closed unclosed code block at line {code_block_start + 1}")
        
        return result
    
    def _fix_tables(self, lines: List[str]) -> List[str]:
        result = []
        i = 0
        
        while i < len(lines):
            line = lines[i]
            
            if self._is_table_row(line):
                table_lines = [line]
                j = i + 1
                
                while j < len(lines) and self._is_table_row(lines[j]):
                    table_lines.append(lines[j])
                    j += 1
                
                table_lines = self._normalize_table(table_lines)
                result.extend(table_lines)
                i = j
            else:
                result.append(line)
                i += 1
        
        return result
    
    def _is_table_row(self, line: str) -> bool:
        stripped = line.strip()
        return '|' in stripped and (stripped.startswith('|') or stripped.endswith('|'))
    
    def _normalize_table_row(self, line: str, num_cols: int) -> str:
        cells = [c.strip() for c in line.split('|') if c.strip()]
        while len(cells) < num_cols:
            cells.append('')
        if len(cells) > num_cols:
            cells = cells[:num_cols]
        return '|' + '|'.join(cells) + '|'
    
    def _normalize_table(self, table_lines: List[str]) -> List[str]:
        if len(table_lines) < 1:
            return table_lines
        
        header_cells = [c.strip() for c in table_lines[0].split('|') if c.strip()]
        num_cols = len(header_cells)
        
        if num_cols == 0:
            return table_lines
        
        has_separator = False
        separator_index = -1
        
        for i, line in enumerate(table_lines[1:], 1):
            cells = [c.strip() for c in line.split('|') if c.strip()]
            if all(c.replace('-', '').replace(':', '') == '' for c in cells if c):
                has_separator = True
                separator_index = i
                break
        
        result = []
        result.append(self._normalize_table_row(table_lines[0], num_cols))
        
        if not has_separator:
            separator = '|' + '|'.join(['---'] * num_cols) + '|'
            result.append(separator)
            self._log_fix(f"Added missing table separator row")
        
        for i, line in enumerate(table_lines):
            if has_separator and i == separator_index:
                normalized_sep = '|' + '|'.join(['---'] * num_cols) + '|'
                result.append(normalized_sep)
                continue
            
            if i == 0:
                continue
            
            if has_separator and i == separator_index:
                continue
            
            normalized_row = self._normalize_table_row(line, num_cols)
            result.append(normalized_row)
        
        return result
    
    def _adjust_heading_levels(self, lines: List[str]) -> List[str]:
        h1_count = 0
        in_code_block = False
        
        for line in lines:
            if line.strip().startswith('```'):
                in_code_block = not in_code_block
                continue
            if not in_code_block and re.match(r'^#\s+', line):
                h1_count += 1
        
        has_multiple_h1 = h1_count > 1
        
        heading_levels = []
        first_h1_found = False
        in_code_block = False
        
        for line in lines:
            if line.strip().startswith('```'):
                in_code_block = not in_code_block
                continue
            if in_code_block:
                continue
            heading_match = re.match(r'^(#{1,6})\s+', line)
            if heading_match:
                level = len(heading_match.group(1))
                if level == 1 and not first_h1_found:
                    first_h1_found = True
                    continue
                if has_multiple_h1:
                    heading_levels.append(level + 1)
                else:
                    heading_levels.append(level)
        
        if not heading_levels:
            return lines
        
        min_level = min(heading_levels)
        
        if has_multiple_h1:
            if min_level <= 2:
                offset = 0
            else:
                offset = min_level - 2
        else:
            if min_level <= 2:
                offset = 0
            else:
                offset = min_level - 2
        
        result = []
        first_h1_processed = False
        in_code_block = False
        
        for line in lines:
            if line.strip().startswith('```'):
                in_code_block = not in_code_block
                result.append(line)
                continue
            
            if in_code_block:
                result.append(line)
                continue
            
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if heading_match:
                level = len(heading_match.group(1))
                text = heading_match.group(2)
                
                if level == 1 and not first_h1_processed:
                    result.append(line)
                    first_h1_processed = True
                    continue
                
                if has_multiple_h1:
                    effective_level = level + 1
                else:
                    effective_level = level
                
                new_level = effective_level - offset
                if new_level < 1:
                    new_level = 1
                if new_level > 6:
                    new_level = 6
                new_line = '#' * new_level + ' ' + text
                result.append(new_line)
                if new_level != level:
                    self._log_fix(f"Adjusted heading level: H{level} -> H{new_level}: '{text[:30]}...'")
            else:
                result.append(line)
        
        return result
    
    def _fix_headings(self, lines: List[str]) -> List[str]:
        result = []
        
        for i, line in enumerate(lines):
            heading_match = re.match(r'^(#{1,6})\s*(.+)$', line)
            if heading_match:
                level = len(heading_match.group(1))
                text = heading_match.group(2).strip()
                text = self._convert_chinese_number_in_heading(text)
                normalized = '#' * level + ' ' + text
                if normalized != line:
                    self._log_fix(f"Normalized heading: '{line.strip()}' -> '{normalized}'")
                result.append(normalized)
            else:
                result.append(line)
        
        return result
    
    def _fix_lists(self, lines: List[str]) -> List[str]:
        result = []
        
        for line in lines:
            stripped = line.strip()
            
            if re.match(r'^[-]{2,}$', stripped) or re.match(r'^[*]{2,}$', stripped) or re.match(r'^[_]{2,}$', stripped):
                result.append(line)
                continue
            
            if stripped.startswith('**'):
                result.append(line)
                continue
            
            bullet_match = re.match(r'^(\s*)[-*+](\s*)(.*)$', line)
            if bullet_match and bullet_match.group(3).strip():
                indent = bullet_match.group(1)
                space = bullet_match.group(2)
                text = bullet_match.group(3).strip()
                normalized = indent + '- ' + text
                if normalized != line:
                    if not space:
                        self._log_fix(f"Added missing space in bullet list: '{line.strip()}'")
                    else:
                        self._log_fix(f"Normalized bullet list: '{line.strip()}'")
                result.append(normalized)
                continue
            
            ordered_match = re.match(r'^(\s*)(\d+)[.)](\s*)(.*)$', line)
            if ordered_match and ordered_match.group(4).strip():
                indent = ordered_match.group(1)
                number = ordered_match.group(2)
                space = ordered_match.group(3)
                text = ordered_match.group(4).strip()
                normalized = f"{indent}{number}. {text}"
                if normalized != line:
                    if not space:
                        self._log_fix(f"Added missing space in ordered list: '{line.strip()}'")
                    else:
                        self._log_fix(f"Normalized ordered list: '{line.strip()}'")
                result.append(normalized)
                continue
            
            result.append(line)
        
        return result
    
    def _fix_horizontal_rules(self, lines: List[str]) -> List[str]:
        result = []
        
        for line in lines:
            stripped = line.strip()
            if re.match(r'^[-_*]{2,}$', stripped) and stripped != '---':
                result.append('---')
                self._log_fix(f"Normalized horizontal rule: '{stripped}' -> '---'")
            else:
                result.append(line)
        
        return result
    
    def _fix_inline_formatting(self, lines: List[str]) -> List[str]:
        result = []
        
        for line in lines:
            if line.strip().startswith('```'):
                result.append(line)
                continue
            
            bold_count = line.count('**')
            if bold_count % 2 != 0:
                line = re.sub(r'\*\*(?!\*)', '', line)
                self._log_fix(f"Removed unmatched bold markers")
            
            italic_count = line.count('*') - line.count('**') * 2
            if italic_count % 2 != 0:
                line = re.sub(r'(?<!\*)\*(?!\*)', '', line)
                self._log_fix(f"Removed unmatched italic markers")
            
            result.append(line)
        
        return result
    
    def _is_ordered_list_item(self, line: str) -> bool:
        stripped = line.strip()
        return bool(re.match(r'^\d+[.)]\s+', stripped))
    
    def _is_heading(self, line: str) -> bool:
        return bool(re.match(r'^#{1,6}\s', line))
    
    def _is_list_item(self, line: str) -> bool:
        stripped = line.strip()
        return bool(re.match(r'^[-*+]\s+', stripped)) or bool(re.match(r'^\d+[.)]\s+', stripped))
    
    def _is_special_line(self, line: str) -> bool:
        if line.strip() == '':
            return False
        return (self._is_heading(line) or 
                line.strip().startswith('>') or
                line.strip().startswith('```') or
                line.strip().startswith('|') or
                line.strip() == '---')
    
    def _fix_empty_lines(self, lines: List[str]) -> List[str]:
        result = []
        prev_empty = False
        empty_count = 0
        i = 0
        
        while i < len(lines):
            line = lines[i]
            
            if line.strip() == '':
                if not prev_empty:
                    if i > 0 and i < len(lines) - 1:
                        prev_line = lines[i - 1]
                        next_line = lines[i + 1]
                        
                        if self._is_heading(prev_line) or self._is_heading(next_line):
                            result.append('')
                            prev_empty = True
                            empty_count = 1
                            i += 1
                            continue
                        
                        if self._is_special_line(prev_line) or self._is_special_line(next_line):
                            result.append('')
                            prev_empty = True
                            empty_count = 1
                            i += 1
                            continue
                        
                        if prev_line.strip() == '---' or next_line.strip() == '---':
                            result.append('')
                            prev_empty = True
                            empty_count = 1
                            i += 1
                            continue
                    
                    self._log_fix(f"Removed empty line between paragraphs")
                    prev_empty = True
                    empty_count = 1
                else:
                    empty_count += 1
            else:
                result.append(line)
                prev_empty = False
            
            i += 1
        
        if empty_count > 2:
            self._log_fix(f"Compressed {empty_count} consecutive empty lines to 1")
        
        return result
    
    def _fix_paragraph_spacing(self, lines: List[str]) -> List[str]:
        result = []
        
        for i, line in enumerate(lines):
            is_heading = bool(re.match(r'^#{1,6}\s', line))
            
            if is_heading and i > 0:
                if result and result[-1].strip() != '':
                    result.append('')
                    self._log_fix(f"Added empty line before heading: '{line.strip()[:30]}...'")
            
            result.append(line)
            
            if is_heading and i < len(lines) - 1:
                if i + 1 < len(lines) and lines[i + 1].strip() != '':
                    pass
        
        return result
    
    def _fix_trailing_whitespace(self, content: str) -> str:
        lines = content.split('\n')
        fixed_lines = []
        trailing_count = 0
        br_converted = 0
        i = 0
        
        while i < len(lines):
            line = lines[i]
            
            if line.endswith('  ') or line.endswith('\t'):
                stripped = line.rstrip()
                merged_line = stripped
                
                while i + 1 < len(lines):
                    next_line = lines[i + 1]
                    if next_line.strip() == '':
                        break
                    if next_line.endswith('  ') or next_line.endswith('\t'):
                        merged_line += '<br>' + next_line.rstrip()
                        br_converted += 1
                        i += 1
                    else:
                        merged_line += '<br>' + next_line.rstrip()
                        br_converted += 1
                        i += 1
                        break
                
                fixed_lines.append(merged_line)
                br_converted += 1
            else:
                stripped = line.rstrip()
                if stripped != line:
                    trailing_count += 1
                fixed_lines.append(stripped)
            
            i += 1
        
        if trailing_count > 0:
            self._log_fix(f"Removed trailing whitespace from {trailing_count} lines")
        
        if br_converted > 0:
            self._log_fix(f"Converted {br_converted} Markdown line breaks to <br>")
        
        return '\n'.join(fixed_lines)
    
    def _fix_leading_whitespace(self, lines: List[str]) -> List[str]:
        result = []
        in_code_block = False
        
        for i, line in enumerate(lines):
            if line.strip().startswith('```'):
                in_code_block = not in_code_block
                result.append(line)
                continue
            
            if in_code_block:
                result.append(line)
                continue
            
            if re.match(r'^#{1,6}\s', line):
                result.append(line)
                continue
            
            if re.match(r'^\s*>', line):
                result.append(line)
                continue
            
            if re.match(r'^\s*[-*+]\s+', line):
                result.append(line)
                continue
            
            if re.match(r'^\s*\d+[.)]\s+', line):
                result.append(line)
                continue
            
            if line.strip() == '':
                result.append(line)
                continue
            
            stripped = line.lstrip()
            if stripped != line:
                self._log_fix(f"Removed leading whitespace from paragraph")
                result.append(stripped)
            else:
                result.append(line)
        
        return result
    
    def get_fixes(self) -> List[str]:
        return self.fixes_applied


def normalize_markdown(content: str, verbose: bool = False) -> Tuple[str, List[str]]:
    normalizer = MarkdownNormalizer(verbose=verbose)
    normalized = normalizer.normalize(content)
    return normalized, normalizer.get_fixes()


def normalize_markdown_file(input_path: str, output_path: Optional[str] = None, verbose: bool = False) -> Tuple[str, List[str]]:
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    normalized, fixes = normalize_markdown(content, verbose=verbose)
    
    if output_path:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(normalized)
        if verbose:
            print(f"[Normalizer] Saved normalized file to: {output_path}")
    
    return normalized, fixes


if __name__ == '__main__':
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python markdown_normalizer.py <input.md> [output.md]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    normalized, fixes = normalize_markdown_file(input_file, output_file, verbose=True)
    
    print(f"\nTotal fixes applied: {len(fixes)}")
