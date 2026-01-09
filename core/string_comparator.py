"""
简要实现文档中提出的字符串比较功能（基础子集）
依赖 textdistance 实现 Levenshtein/Jaro-Winkler 等
"""
import re
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







