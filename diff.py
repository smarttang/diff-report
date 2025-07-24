#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自动分析 MT4/MT5 回测报告（xlsx）目录
Author: Smarttang
"""

import pandas as pd
import sys
import os
import glob
from collections import defaultdict

def process_file(xlsx_path: str):
    """处理单个回测文件并返回分析结果"""
    # 1. 读取全部数据
    df = pd.read_excel(xlsx_path)
    
    # 2. 去掉最后一行（统计行）
    df = df.iloc[:-1]
    
    # 3. 检查必要列
    required_cols = {'交易品种', '盈利'}
    missing = required_cols - set(df.columns)
    if missing:
        raise ValueError(f"文件 {os.path.basename(xlsx_path)} 缺少必要列：{missing}")
    
    # 4. 汇总：盈利、亏损、净盈亏、交易次数
    summary = (
        df.groupby('交易品种')['盈利']
          .agg(total_profit=lambda s: s[s > 0].sum(),
               total_loss=lambda s: s[s < 0].sum(),
               net_profit='sum',
               trades='count')
          .sort_values('net_profit', ascending=False)
    )
    
    # 5. 提取最佳和最差品种
    top_profit_symbol = summary['net_profit'].idxmax()
    top_profit_value = summary.loc[top_profit_symbol, 'net_profit']
    
    loss_df = summary[summary['net_profit'] < 0]
    if not loss_df.empty:
        top_loss_symbol = loss_df['net_profit'].idxmin()
        top_loss_value = summary.loc[top_loss_symbol, 'net_profit']
    else:
        top_loss_symbol = None
        top_loss_value = 0
    
    return {
        'filename': os.path.basename(xlsx_path),
        'summary': summary,
        'top_profit': (top_profit_symbol, top_profit_value),
        'top_loss': (top_loss_symbol, top_loss_value) if top_loss_symbol else None
    }

def analyze_directory(directory: str):
    """分析目录中的所有回测文件"""
    # 1. 获取所有xlsx文件
    files = glob.glob(os.path.join(directory, '*.xlsx'))
    if not files:
        print(f"目录 {directory} 中没有找到xlsx文件")
        return
    
    # 2. 处理所有文件
    all_results = []
    symbol_records = defaultdict(list)  # 记录每个品种在所有文件中的表现
    
    print(f"========== 开始分析目录: {directory} ==========")
    print(f"找到 {len(files)} 个回测文件\n")
    
    for file_path in files:
        try:
            result = process_file(file_path)
            all_results.append(result)
            
            # 输出单个文件报告
            print(f"\n★ 文件: {result['filename']}")
            print("=" * 50)
            print(result['summary'].to_string(float_format="%.2f"))
            
            # 记录品种表现
            for symbol, row in result['summary'].iterrows():
                symbol_records[symbol].append(row['net_profit'])
            
            # 输出最佳/最差品种
            profit_symbol, profit_value = result['top_profit']
            print(f"\n最佳品种: {profit_symbol} 净盈利 {profit_value:.2f}")
            
            if result['top_loss']:
                loss_symbol, loss_value = result['top_loss']
                print(f"最差品种: {loss_symbol} 净亏损 {-loss_value:.2f}")
            else:
                print("该文件没有亏损品种")
                
            print("=" * 50)
            
        except Exception as e:
            print(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}")
    
    # 3. 全局汇总分析
    print("\n" + "=" * 60)
    print(" " * 20 + "全局汇总分析")
    print("=" * 60)
    
    # 3.1 合并所有品种数据
    global_summary = pd.concat([res['summary'] for res in all_results])
    global_summary = global_summary.groupby(global_summary.index).sum()
    global_summary = global_summary.sort_values('net_profit', ascending=False)
    
    print("\n各品种全局表现:")
    print(global_summary.to_string(float_format="%.2f"))
    
    # 3.2 全局最佳/最差品种
    global_top_profit = global_summary['net_profit'].idxmax()
    global_top_loss = global_summary['net_profit'].idxmin()
    
    print(f"\n全局最佳品种: {global_top_profit} "
          f"净盈利 {global_summary.loc[global_top_profit, 'net_profit']:.2f}")
    
    print(f"全局最差品种: {global_top_loss} "
          f"净亏损 {-global_summary.loc[global_top_loss, 'net_profit']:.2f}")
    
    # 3.3 稳定盈利和持续亏损品种分析
    profitable_symbols = []
    loss_symbols = []
    
    for symbol, profits in symbol_records.items():
        # 至少出现在2个文件中才考虑稳定性
        if len(profits) < 2:
            continue
            
        if all(p > 0 for p in profits):
            avg_profit = sum(profits) / len(profits)
            profitable_symbols.append((symbol, avg_profit, len(profits)))
        
        elif all(p < 0 for p in profits):
            avg_loss = sum(profits) / len(profits)
            loss_symbols.append((symbol, avg_loss, len(profits)))
    
    # 按盈利/亏损程度排序
    profitable_symbols.sort(key=lambda x: x[1], reverse=True)
    loss_symbols.sort(key=lambda x: x[1])
    
    # 输出稳定盈利品种
    if profitable_symbols:
        print("\n稳定盈利品种:")
        for symbol, avg_profit, count in profitable_symbols:
            print(f"{symbol}: 平均盈利 {avg_profit:.2f} (在 {count} 个文件中持续盈利)")
    else:
        print("\n没有稳定盈利的品种")
    
    # 输出持续亏损品种
    if loss_symbols:
        print("\n持续亏损品种:")
        for symbol, avg_loss, count in loss_symbols:
            print(f"{symbol}: 平均亏损 {avg_loss:.2f} (在 {count} 个文件中持续亏损)")
    else:
        print("\n没有持续亏损的品种")

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("用法：python analyze_backtest.py <目录路径>")
        sys.exit(1)
    
    directory = sys.argv[1]
    if not os.path.isdir(directory):
        print(f"错误: {directory} 不是有效目录")
        sys.exit(1)
    
    analyze_directory(directory)
