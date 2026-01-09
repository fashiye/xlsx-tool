"""
简要实现文档中提出的字符串比较功能（基础子集）
依赖 textdistance 实现 Levenshtein/Jaro-Winkler 等
"""
import difflib
import re
import textdistance
import logging

# 配置日志记录
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("app.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

class StringComparator:
    def exact_match(self, s1, s2, ignore_case=False):
        logger.info(f"执行精确字符串匹配: '{s1}' vs '{s2}', 忽略大小写: {ignore_case}")
        if ignore_case:
            result = str(s1).lower() == str(s2).lower()
        else:
            result = str(s1) == str(s2)
        logger.info(f"精确匹配结果: {result}")
        return result

    def fuzzy_match(self, s1, s2, method='levenshtein', threshold=0.8):
        logger.info(f"执行模糊字符串匹配: '{s1}' vs '{s2}', 方法: {method}, 阈值: {threshold}")
        s1 = str(s1) or ""
        s2 = str(s2) or ""
        try:
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
            result = (d >= threshold, d)
            logger.info(f"模糊匹配结果: 匹配 = {result[0]}, 相似度 = {result[1]}")
            return result
        except Exception as e:
            logger.error(f"模糊匹配失败: {str(e)}")
            return (False, 0.0)

    def regex_match(self, s1, s2, pattern):
        logger.info(f"执行正则表达式匹配: '{s1}' vs '{s2}', 模式: {pattern}")
        try:
            m1 = re.search(pattern, s1)
            m2 = re.search(pattern, s2)
            result = bool(m1 and m2)
            logger.info(f"正则匹配结果: {result}")
            return result
        except re.error as e:
            logger.error(f"正则表达式错误: {str(e)}")
            return False

    def text_diff(self, s1, s2, mode='line'):
        logger.info(f"执行文本差异分析: '{s1[:50]}...' vs '{s2[:50]}...', 模式: {mode}")
        try:
            if mode == 'line':
                lines1 = str(s1).splitlines()
                lines2 = str(s2).splitlines()
                diff = list(difflib.unified_diff(lines1, lines2, lineterm=''))
                has_diff = any(not (line.startswith('---') or line.startswith('+++') or line.startswith('@@')) and (line.startswith('+') or line.startswith('-')) for line in diff)
                result = (diff, has_diff)
            else:
                diff = list(difflib.ndiff(str(s1), str(s2)))
                has_diff = any(ch[0] in ('+', '-') for ch in diff)
                result = (diff, has_diff)
            logger.info(f"文本差异分析结果: 存在差异 = {result[1]}")
            return result
        except Exception as e:
            logger.error(f"文本差异分析失败: {str(e)}")
            return ([], True)

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