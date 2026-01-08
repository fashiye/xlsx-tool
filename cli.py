#!/usr/bin/env python3
"""
Excel数据对比工具 - 命令行接口

提供脱离GUI的独立运行能力，实现所有核心业务功能
"""
import argparse
import logging
import sys
from core.comparison_service import ComparisonService

# 配置日志记录
logging.basicConfig(level=logging.DEBUG, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("app.log"),
                        logging.StreamHandler()
                    ])
logger = logging.getLogger(__name__)

def main():
    """
    命令行接口主函数
    """
    parser = argparse.ArgumentParser(description='Excel数据对比工具')
    parser.add_argument('--file1', '-f1', default='test_logging.xlsx', help='第一个Excel文件路径')
    parser.add_argument('--file2', '-f2', default='test_logging.xlsx', help='第二个Excel文件路径')
    parser.add_argument('--sheet1', '-s1', default='Sheet1', help='第一个文件的工作表名称（可选）')
    parser.add_argument('--sheet2', '-s2', default='Sheet1', help='第二个文件的工作表名称（可选）')
    parser.add_argument('--rule', '-r', action='append', help='比较规则，可以多次使用（可选）')
    parser.add_argument('--output', '-o', default=None, help='结果输出文件路径（可选）')
    parser.add_argument('--tolerance', '-t', default='0', help='数值比较容差（可选）')
    parser.add_argument('--ignore-case', action='store_true', help='字符串比较忽略大小写（可选）')
    
    args = parser.parse_args()
    
    print("=" * 60)
    print("Excel数据对比工具 - 命令行模式")
    print("=" * 60)
    
#    try:
    if True:
    
        # 创建比较服务实例
        service = ComparisonService()
        
        # 加载第一个文件
        print(f"\n正在加载文件1: {args.file1}")
        sheets1 = service.load_workbook(args.file1, "file1")
        sheet1 = args.sheet1 or sheets1[0]
        service.load_sheet_data("file1", sheet1)
        print(f"  ✓ 文件1加载成功，使用工作表: {sheet1}")
        
        # 加载第二个文件
        print(f"\n正在加载文件2: {args.file2}")
        sheets2 = service.load_workbook(args.file2, "file2")
        sheet2 = args.sheet2 or sheets2[0]
        service.load_sheet_data("file2", sheet2)
        print(f"  ✓ 文件2加载成功，使用工作表: {sheet2}")
        
        # 添加规则
        if args.rule:
            print(f"\n正在添加{len(args.rule)}条比较规则")
            for i, rule in enumerate(args.rule, 1):
                try:
                    service.add_rule(rule)
                    print(f"  ✓ 规则{i}: {rule}")
                except Exception as e:
                    print(f"  ✗ 规则{i}: {rule} - 错误: {e}")
                    logger.error(f"添加规则失败: {rule}，错误: {str(e)}")
        
        # 执行比较
        print("\n正在执行比较...")
        use_rules = args.rule is not None and len(args.rule) > 0
        options = {
            'tolerance': args.tolerance,
            'ignore_case': args.ignore_case
        }
        
        result_text, result_df, result_map = service.run_comparison(use_rules=use_rules, options=options)
        
        # 打印结果
        print("\n" + "=" * 60)
        print("比较结果")
        print("=" * 60)
        print(result_text)
        
        # 保存结果
        if args.output:
            print(f"\n正在保存结果到: {args.output}")
            try:
                service.save_results(result_df, args.output)
                print(f"  ✓ 结果保存成功")
            except Exception as e:
                print(f"  ✗ 结果保存失败: {e}")
                logger.error(f"保存结果失败: {str(e)}")
        
        print("\n" + "=" * 60)
        print("比较完成")
        print("=" * 60)
        
#    except Exception as e:
 #       print(f"\n错误: {e}")
  #      logger.error(f"比较失败: {str(e)}")
   #     sys.exit(1)

if __name__ == "__main__":
    main()
