"""
简要实现文档中提出的字符串比较功能（基础子集）
依赖 textdistance 实现 Levenshtein/Jaro-Winkler 等
"""
import difflib
import re
import textdistance

class StringComparator:
    def exact_match(self, s1, s2, ignore_case=False):
        if ignore_case:
            return str(s1).lower() == str(s2).lower()
        return str(s1) == str(s2)

    def fuzzy_match(self, s1, s2, method='levenshtein', threshold=0.8):
        s1 = str(s1) or ""
        s2 = str(s2) or ""
        if method in ('levenshtein', 'lev'):
            if textdistance:
                d = textdistance.levenshtein.normalized_similarity(s1, s2)
            else:
                # fallback: ratio from difflib
                d = difflib.SequenceMatcher(None, s1, s2).ratio()
        elif method in ('jaro', 'jaro_winkler', 'jw'):
            if textdistance:
                d = textdistance.jaro_winkler.normalized_similarity(s1, s2)
            else:
                d = difflib.SequenceMatcher(None, s1, s2).ratio()
        else:
            d = difflib.SequenceMatcher(None, s1, s2).ratio()
        return (d >= threshold, d)

    def regex_match(self, s1, s2, pattern):
        try:
            m1 = re.search(pattern, s1)
            m2 = re.search(pattern, s2)
            return bool(m1 and m2)
        except re.error:
            return False

    def text_diff(self, s1, s2, mode='line'):
        if mode == 'line':
            lines1 = str(s1).splitlines()
            lines2 = str(s2).splitlines()
            diff = list(difflib.unified_diff(lines1, lines2, lineterm=''))
            has_diff = any(not (line.startswith('---') or line.startswith('+++') or line.startswith('@@')) and (line.startswith('+') or line.startswith('-')) for line in diff)
            return diff, has_diff
        else:
            diff = list(difflib.ndiff(str(s1), str(s2)))
            has_diff = any(ch[0] in ('+', '-') for ch in diff)
            return diff, has_diff

    def structured_match(self, s1, s2, format_type='csv', delimiter=','):
        if format_type == 'csv':
            l1 = [x.strip() for x in str(s1).split(delimiter)]
            l2 = [x.strip() for x in str(s2).split(delimiter)]
            equal = l1 == l2
            return equal, {'left': l1, 'right': l2}
        elif format_type == 'json':
            import json
            try:
                o1 = json.loads(s1)
                o2 = json.loads(s2)
                return o1 == o2, {'left': o1, 'right': o2}
            except:
                return False, "JSON解析失败"
        return False, {}