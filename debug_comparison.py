#!/usr/bin/env python3
"""
调试比较逻辑的简单脚本
"""
import logging

# 配置日志
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# 测试比较逻辑
def test_comparison():
    # 模拟计算结果
    left_value = 2.0  # 浮点数
    right_value = 2    # 整数
    op = "="
    
    print(f"左值: {left_value} (类型: {type(left_value)})")
    print(f"右值: {right_value} (类型: {type(right_value)})")
    print(f"操作符: {op}")
    
    # 执行比较
    try:
        left_scalar = float(left_value)
        right_scalar = float(right_value)
        
        print(f"转换后左值: {left_scalar} (类型: {type(left_scalar)})")
        print(f"转换后右值: {right_scalar} (类型: {type(right_scalar)})")
        
        if op == '=':
            result = abs(float(left_scalar) - float(right_scalar)) < 1e-6
        
        print(f"比较结果: {result} (类型: {type(result)})")
        
        # 手动计算差异
        diff = abs(left_scalar - right_scalar)
        print(f"差异值: {diff}")
        print(f"是否小于阈值(1e-6): {diff < 1e-6}")
        
    except Exception as e:
        print(f"比较错误: {e}")

if __name__ == "__main__":
    test_comparison()