import pandas as pd
import logging
import tempfile
import os
from core.comparator import ExcelComparator

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_test_excel(filepath, sheets=None):
    """
    创建测试用的Excel文件
    
    参数:
        filepath: 输出文件路径
        sheets: 字典，键为工作表名，值为DataFrame
    """
    sheets = sheets or {}
    with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

def test_nonexistent_workbook():
    """
    测试访问不存在的工作簿
    """
    logger.info("开始测试访问不存在的工作簿")
    comparator = ExcelComparator()
    
    try:
        comparator.get_sheet_dataframe("nonexistent", "Sheet1")
        assert False, "应该抛出ValueError"
    except ValueError as e:
        assert "未加载工作簿" in str(e)
        logger.info("访问不存在的工作簿测试通过")

def test_nonexistent_sheet():
    """
    测试访问不存在的工作表
    """
    logger.info("开始测试访问不存在的工作表")
    comparator = ExcelComparator()
    
    # 创建测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # 创建包含单个工作表的测试文件
        sheets = {"Sheet1": pd.DataFrame({"A": [1, 2, 3]})}
        create_test_excel(tmp_path, sheets)
        
        # 加载工作簿
        comparator.load_workbook(tmp_path, alias="TEST")
        
        # 尝试访问不存在的工作表
        try:
            comparator.get_sheet_dataframe("TEST", "Sheet2")
            assert False, "应该抛出ValueError"
        except ValueError as e:
            assert "工作表 Sheet2 不存在" in str(e)
            logger.info("访问不存在的工作表测试通过")
    finally:
        # 清理临时文件
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

def test_empty_sheet():
    """
    测试空工作表的处理
    """
    logger.info("开始测试空工作表的处理")
    comparator = ExcelComparator()
    
    # 创建测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # 创建包含空工作表的测试文件
        sheets = {"EmptySheet": pd.DataFrame()}
        create_test_excel(tmp_path, sheets)
        
        # 加载工作簿
        comparator.load_workbook(tmp_path, alias="TEST")
        
        # 获取空工作表
        df = comparator.get_sheet_dataframe("TEST", "EmptySheet")
        assert df.shape == (0, 0), f"空工作表形状应该是 (0, 0)，实际是 {df.shape}"
        logger.info("空工作表处理测试通过")
    finally:
        # 清理临时文件
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

def test_compare_with_rules_nonexistent_sheet():
    """
    测试使用不存在的工作表进行规则比较
    """
    logger.info("开始测试使用不存在的工作表进行规则比较")
    comparator = ExcelComparator()
    
    # 创建测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # 创建测试文件
        sheets = {"Sheet1": pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})}
        create_test_excel(tmp_path, sheets)
        
        # 加载工作簿
        comparator.load_workbook(tmp_path, alias="TEST")
        
        # 添加规则
        comparator.add_rule("A1 + B1 = C1")
        
        # 尝试使用不存在的工作表进行比较
        result_summary, comparison_results = comparator.compare_with_rules("TEST", "NonexistentSheet")
        
        # 验证结果
        assert result_summary['error'] is not None
        assert "工作表 NonexistentSheet 不存在" in result_summary['error']
        logger.info("使用不存在的工作表进行规则比较测试通过")
    finally:
        # 清理临时文件和规则
        comparator.clear_rules()
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

def test_compare_sheets_with_rules_invalid_input():
    """
    测试compare_sheets_with_rules方法的无效输入
    """
    logger.info("开始测试compare_sheets_with_rules方法的无效输入")
    comparator = ExcelComparator()
    
    # 创建两个测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp1, \
         tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp2:
        tmp_path1 = tmp1.name
        tmp_path2 = tmp2.name
    
    try:
        # 创建测试文件
        sheets1 = {"Sheet1": pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})}
        sheets2 = {"Sheet1": pd.DataFrame({"A": [1, 2, 3], "B": [4, 5, 6]})}
        create_test_excel(tmp_path1, sheets1)
        create_test_excel(tmp_path2, sheets2)
        
        # 加载工作簿
        comparator.load_workbook(tmp_path1, alias="FILE1")
        comparator.load_workbook(tmp_path2, alias="FILE2")
        
        # 添加规则
        comparator.add_rule("A1 = FILE2:A1")
        
        # 测试使用不存在的工作表1
        result_summary, comparison_results, combined_df = comparator.compare_sheets_with_rules("FILE1", "NonexistentSheet", "FILE2", "Sheet1")
        assert result_summary['error'] is not None
        assert "工作表 NonexistentSheet 不存在" in result_summary['error']
        
        # 测试使用不存在的工作表2
        result_summary, comparison_results, combined_df = comparator.compare_sheets_with_rules("FILE1", "Sheet1", "FILE2", "NonexistentSheet")
        assert result_summary['error'] is not None
        assert "工作表 NonexistentSheet 不存在" in result_summary['error']
        
        logger.info("compare_sheets_with_rules方法无效输入测试通过")
    finally:
        # 清理临时文件和规则
        comparator.clear_rules()
        if os.path.exists(tmp_path1):
            os.unlink(tmp_path1)
        if os.path.exists(tmp_path2):
            os.unlink(tmp_path2)

def test_boundary_conditions():
    """
    测试边界条件，如空值、超大数值等
    """
    logger.info("开始测试边界条件")
    comparator = ExcelComparator()
    
    # 创建测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # 创建包含边界条件的测试文件
        sheets = {"BoundarySheet": pd.DataFrame({
            "A": [None, 0, 1, -1, 10**10, -10**10, 1.123456789123456789],
            "B": [None, 0, 2, -2, 10**10 + 1, -10**10 - 1, 1.123456789123456788]
        })}
        create_test_excel(tmp_path, sheets)
        
        # 加载工作簿
        comparator.load_workbook(tmp_path, alias="TEST")
        
        # 获取数据
        df = comparator.get_sheet_dataframe("TEST", "BoundarySheet")
        
        # 测试空值比较
        comparator.add_rule("A1 = B1")  # 两个空值应该相等
        result_summary, comparison_results = comparator.compare_with_rules("TEST", "BoundarySheet")
        assert len(comparison_results['passed']) == 1  # 空值比较应该通过
        
        # 测试超大数值比较
        comparator.clear_rules()
        comparator.add_rule("A5 = B5")  # 超大数值，应该不相等
        result_summary, comparison_results = comparator.compare_with_rules("TEST", "BoundarySheet")
        assert len(comparison_results['failed']) == 1  # 超大数值比较应该失败
        
        # 测试浮点精度
        comparator.clear_rules()
        comparator.add_rule("A7 = B7")  # 浮点精度测试，应该通过
        result_summary, comparison_results = comparator.compare_with_rules("TEST", "BoundarySheet")
        assert len(comparison_results['passed']) == 1  # 浮点精度测试应该通过（因为有容差）
        
        logger.info("边界条件测试通过")
    finally:
        # 清理临时文件和规则
        comparator.clear_rules()
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

def test_select_cells_invalid_range():
    """
    测试选择无效单元格范围的情况
    """
    logger.info("开始测试选择无效单元格范围")
    comparator = ExcelComparator()
    
    # 创建测试文件
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp:
        tmp_path = tmp.name
    
    try:
        # 创建测试文件
        sheets = {"Sheet1": pd.DataFrame({
            "A": [1, 2, 3],
            "B": [4, 5, 6],
            "C": [7, 8, 9]
        })}
        create_test_excel(tmp_path, sheets)
        
        # 加载工作簿
        comparator.load_workbook(tmp_path, alias="TEST")
        
        # 测试超出范围的选择
        df = comparator.select_cells("TEST", "Sheet1", "A1:Z100")
        assert df.shape == (3, 3)  # 应该返回整个工作表的数据
        
        # 测试单个单元格选择
        df = comparator.select_cells("TEST", "Sheet1", "B2")
        assert df.shape == (1, 1)  # 应该返回单个单元格
        assert df.iloc[0, 0] == 5  # 单元格值应该是5
        
        logger.info("选择无效单元格范围测试通过")
    finally:
        # 清理临时文件
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

if __name__ == "__main__":
    # 运行所有测试
    test_nonexistent_workbook()
    test_nonexistent_sheet()
    test_empty_sheet()
    test_compare_with_rules_nonexistent_sheet()
    test_compare_sheets_with_rules_invalid_input()
    test_boundary_conditions()
    test_select_cells_invalid_range()
    print("所有工作表输入错误和边界条件测试通过！")