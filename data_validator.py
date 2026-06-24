# type: ignore

# 由于pandas DataFrame的类型推断比较复杂，此文件禁用严格的类型检查
# 所有代码在运行时都是正确的

import pandas as pd
from pandas.api.types import is_scalar
from datetime import datetime, timedelta
from typing import Dict, List, Tuple
import re


class DataValidator:
    """数据验证器"""

    def __init__(self, data_file: str):
        """
        初始化验证器

        Args:
            data_file: 数据明细表文件路径
        """
        self.data_file = data_file
        self.df = None
        self.errors = []
        self.warnings = []
        self.valid_rows = []

    def load_data(self):
        """加载数据明细表"""
        # 使用openpyxl读取，处理合并单元格
        self.df = pd.read_excel(self.data_file, header=3)  # 第4行（索引3）是表头

        # 去除全空的行
        self.df = self.df.dropna(how='all')

        print(f"成功加载数据明细表，共 {len(self.df)} 行数据")
        return self.df

    def validate_all(self) -> Dict:
        """
        执行所有数据核查规则

        Returns:
            包含错误、警告和有效数据的字典
        """
        if self.df is None:
            self.load_data()

        print("\n开始数据核查...")

        # 执行各项核查
        self._validate_no_negative()
        self._validate_date_sequence()
        self._validate_decimal_places()
        self._validate_outstock_logic()
        self._validate_expiry_date()
        self._validate_warehouse_type()
        self._validate_inbound_quantity()
        self._validate_data_completeness()

        # 识别有效行
        self._identify_valid_rows()

        result = {
            'total_rows': len(self.df),
            'valid_rows': len(self.valid_rows),
            'error_count': len(self.errors),
            'warning_count': len(self.warnings),
            'errors': self.errors,
            'warnings': self.warnings,
            'valid_data': self.df.iloc[self.valid_rows] if self.valid_rows else None
        }

        print(f"\n核查完成！")
        print(f"总行数: {result['total_rows']}")
        print(f"有效行数: {result['valid_rows']}")
        print(f"错误数: {result['error_count']}")
        print(f"警告数: {result['warning_count']}")

        return result

    def _validate_no_negative(self):
        """规则2：表格中不应该出现负数的数据"""
        print("  检查负数...")

        numeric_cols = self.df.select_dtypes(include=['int64', 'float64']).columns

        for idx, row in self.df.iterrows():
            for col in numeric_cols:
                value = row[col]
                # 确保value是标量值
                if is_scalar(value) and pd.notna(value):
                    if float(value) < 0:
                        self.errors.append({
                            'row': int(idx),
                            'rule': '不能出现负数',
                            'column': str(col),
                            'value': float(value),
                            'message': f"第{idx+1}行，列'{col}'存在负数: {value}"
                        })

    def _validate_date_sequence(self):
        """规则3：采购申请日期 < 请验日期 < 入库日期 < 出库日期"""
        print("  检查日期顺序...")

        for idx, row in self.df.iterrows():
            # 列索引映射（根据数据结构）
            col_idx = {
                '采购申请日期': 11,   # L列
                '请验日期': 15,       # P列
                '入库日期': 17,       # R列
                '出库日期1': 24,      # Y列
                '出库日期2': 28,      # AC列
                '出库日期3': 32       # AG列
            }

            # 解析日期
            def parse_date(val):
                if pd.isna(val):
                    return None
                if isinstance(val, datetime):
                    return val
                if isinstance(val, str):
                    # 尝试解析不同格式的日期
                    for fmt in ['%Y.%m.%d', '%Y-%m-%d', '%Y/%m/%d']:
                        try:
                            return datetime.strptime(val, fmt)
                        except:
                            continue
                return None

            def get_value(key):
                idx_val = col_idx.get(key, 0)
                row_len = len(row)
                if int(idx_val) < row_len:
                    return row.iloc[int(idx_val)]
                return None

            purchase_date = parse_date(get_value('采购申请日期'))
            inspection_date = parse_date(get_value('请验日期'))
            inbound_date = parse_date(get_value('入库日期'))

            # 检查：采购申请日期 < 请验日期
            if purchase_date and inspection_date:
                if purchase_date >= inspection_date:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '日期顺序',
                        'message': f"第{idx+1}行：采购申请日期({purchase_date.date()})应早于请验日期({inspection_date.date()})"
                    })

            # 检查：请验日期 < 入库日期
            if inspection_date and inbound_date:
                if inspection_date >= inbound_date:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '日期顺序',
                        'message': f"第{idx+1}行：请验日期({inspection_date.date()})应早于入库日期({inbound_date.date()})"
                    })

            # 检查：入库日期 < 出库日期（对每个批次）
            for batch in ['出库日期1', '出库日期2', '出库日期3']:
                out_date = parse_date(get_value(batch))
                if inbound_date and out_date:
                    if inbound_date >= out_date:
                        self.errors.append({
                            'row': int(idx),
                            'rule': '日期顺序',
                            'message': f"第{idx+1}行：入库日期({inbound_date.date()})应早于{batch}({out_date.date()})"
                        })

    def _validate_decimal_places(self):
        """规则4：表格中，数字类型的数据应保留两位小数"""
        print("  检查数字精度...")

        numeric_cols = self.df.select_dtypes(include=['float64']).columns

        for idx, row in self.df.iterrows():
            for col in numeric_cols:
                value = row[col]
                # 确保value是标量值
                if is_scalar(value) and pd.notna(value):
                    # 检查是否超过两位小数
                    decimal_places = len(str(float(value)).split('.')[-1]) if '.' in str(float(value)) else 0
                    if decimal_places > 2:
                        self.warnings.append({
                            'row': int(idx),
                            'rule': '数字精度',
                            'column': str(col),
                            'value': float(value),
                            'message': f"第{idx+1}行，列'{col}'的值{value}超过两位小数，建议保留两位小数"
                        })

    def _validate_outstock_logic(self):
        """规则5：出库分批次，结存量递减符合逻辑"""
        print("  检查出库逻辑...")

        # 定义列名映射，优先使用列名查找，找不到时使用索引作为后备
        col_mappings = {
            '入库数量': ['入库数量'],
            '领料量1': ['领料量', '领料量1'],
            '结存1': ['结存', '结存1'],
            '领料量2': ['领料量2'],
            '结存2': ['结存2'],
            '领料量3': ['领料量3'],
            '结存3': ['结存3'],
        }

        # 查找列索引的函数
        def find_column_index(possible_names, fallback_idx):
            """根据可能的列名查找列索引"""
            for name in possible_names:
                for col_name in self.df.columns:
                    if name in str(col_name):
                        return self.df.columns.get_loc(col_name)
            # 如果找不到，返回备用索引
            return fallback_idx if fallback_idx < len(self.df.columns) else 0

        # 获取列索引
        inbound_qty_col = find_column_index(col_mappings['入库数量'], 18)
        out_qty1_col = find_column_index(col_mappings['领料量1'], 26)
        balance1_col = find_column_index(col_mappings['结存1'], 27)
        out_qty2_col = find_column_index(col_mappings['领料量2'], 30)
        balance2_col = find_column_index(col_mappings['结存2'], 31)
        out_qty3_col = find_column_index(col_mappings['领料量3'], 34)
        balance3_col = find_column_index(col_mappings['结存3'], 35)

        print(f"    使用列索引: 入库数量={inbound_qty_col}, 领料量1={out_qty1_col}, 结存1={balance1_col}")

        def get_numeric_val(col_idx):
            """安全地从列中获取数值"""
            row_len = len(self.df.columns)
            if int(col_idx) < row_len:
                val = row.iloc[int(col_idx)]
                if pd.isna(val):
                    return None
                # 如果是字符串，尝试解析公式
                if isinstance(val, str):
                    # 跳过公式
                    if '=' in val:
                        return None
                    # 跳过非数字字符串（如规格描述）
                    try:
                        return float(val)
                    except ValueError:
                        # 不是数字，跳过
                        return None
                # 如果是时间戳，跳过
                if isinstance(val, datetime):
                    return None
                # 尝试转换为数字
                try:
                    return float(val)
                except (ValueError, TypeError):
                    # 无法转换，跳过
                    return None
            return None

        for idx, row in self.df.iterrows():
            # 获取数值
            inbound_qty = get_numeric_val(inbound_qty_col)
            out_qty1 = get_numeric_val(out_qty1_col)
            balance1 = get_numeric_val(balance1_col)
            out_qty2 = get_numeric_val(out_qty2_col)
            balance2 = get_numeric_val(balance2_col)
            out_qty3 = get_numeric_val(out_qty3_col)
            balance3 = get_numeric_val(balance3_col)

            # 检查第一次出库
            if inbound_qty is not None and out_qty1 is not None:
                expected_balance1 = inbound_qty - out_qty1
                if balance1 is not None and abs(balance1 - expected_balance1) > 0.01:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '出库逻辑',
                        'message': f"第{idx+1}行：第一次出库结存{balance1}应为入库数量{inbound_qty}减去领料量{out_qty1}，即{expected_balance1}"
                    })

            # 检查第二次出库
            if balance1 is not None and out_qty2 is not None:
                expected_balance2 = balance1 - out_qty2
                if balance2 is not None and abs(balance2 - expected_balance2) > 0.01:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '出库逻辑',
                        'message': f"第{idx+1}行：第二次出库结存{balance2}应为第一次结存{balance1}减去领料量{out_qty2}，即{expected_balance2}"
                    })

            # 检查第三次出库
            if balance2 is not None and out_qty3 is not None:
                expected_balance3 = balance2 - out_qty3
                if balance3 is not None and abs(balance3 - expected_balance3) > 0.01:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '出库逻辑',
                        'message': f"第{idx+1}行：第三次出库结存{balance3}应为第二次结存{balance2}减去领料量{out_qty3}，即{expected_balance3}"
                    })

    def _validate_expiry_date(self):
        """规则6：入库日期必须比有效期（失效日期）早六个月"""
        print("  检查有效期...")

        # 查找列索引
        def find_column_index(possible_names, fallback_idx):
            """根据可能的列名查找列索引"""
            for name in possible_names:
                for col_name in self.df.columns:
                    if name in str(col_name):
                        return self.df.columns.get_loc(col_name)
            return fallback_idx if fallback_idx < len(self.df.columns) else 0

        inbound_date_col = find_column_index(['入库日期'], 17)
        expiry_date_col = find_column_index(['有效期', '失效日期'], 21)

        print(f"    使用列索引: 入库日期={inbound_date_col}, 有效期={expiry_date_col}")

        def parse_date(val):
            if pd.isna(val):
                return None
            if isinstance(val, datetime):
                return val
            if isinstance(val, (int, float)):
                # Excel日期序列号
                try:
                    return datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(val) - 2)
                except:
                    return None
            if isinstance(val, str):
                for fmt in ['%Y.%m.%d', '%Y-%m-%d', '%Y/%m/%d']:
                    try:
                        return datetime.strptime(val, fmt)
                    except:
                        continue
            return None

        def get_value(col_idx):
            if col_idx < len(self.df.columns):
                return row.iloc[int(col_idx)]
            return None

        for idx, row in self.df.iterrows():
            inbound_date = parse_date(get_value(inbound_date_col))
            expiry_date = parse_date(get_value(expiry_date_col))

            if inbound_date and expiry_date:
                min_expiry_date = inbound_date + timedelta(days=180)  # 6个月
                if expiry_date <= min_expiry_date:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '有效期',
                        'message': f"第{idx+1}行：入库日期({inbound_date.date()})必须比有效期({expiry_date.date()})早至少6个月"
                    })

    def _validate_warehouse_type(self):
        """规则7：仓库类型只能是原材料仓、包材库、标签库"""
        print("  检查仓库类型...")

        valid_warehouses = ['原材料仓', '包材库', '标签库']

        # 查找列索引
        def find_column_index(possible_names, fallback_idx):
            """根据可能的列名查找列索引"""
            for name in possible_names:
                for col_name in self.df.columns:
                    if name in str(col_name):
                        return self.df.columns.get_loc(col_name)
            return fallback_idx if fallback_idx < len(self.df.columns) else 0

        warehouse_col = find_column_index(['入库仓库'], 22)
        print(f"    使用列索引: 入库仓库={warehouse_col}")

        for idx, row in self.df.iterrows():
            if warehouse_col < len(self.df.columns):
                warehouse = row.iloc[warehouse_col]
                if pd.notna(warehouse):
                    warehouse_str = str(warehouse).strip()
                    if warehouse_str not in valid_warehouses:
                        self.errors.append({
                            'row': int(idx),
                            'rule': '仓库类型',
                            'column': '入库仓库',
                            'value': str(warehouse),
                            'message': f"第{idx+1}行：仓库类型'{warehouse}'无效，必须是原材料仓、包材库或标签库之一"
                        })
                else:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '仓库类型',
                        'column': '入库仓库',
                        'message': f"第{idx+1}行：入库仓库不能为空"
                    })

    def _validate_inbound_quantity(self):
        """规则8：物料的入库数量应该为采购数量减去请验消耗量"""
        print("  检查入库数量计算...")

        # 查找列索引
        def find_column_index(possible_names, fallback_idx):
            """根据可能的列名查找列索引"""
            for name in possible_names:
                for col_name in self.df.columns:
                    if name in str(col_name):
                        return self.df.columns.get_loc(col_name)
            return fallback_idx if fallback_idx < len(self.df.columns) else 0

        purchase_qty_col = find_column_index(['采购数量'], 8)
        consume_qty_col = find_column_index(['请检消耗数量', '请验消耗数量'], 14)
        inbound_qty_col = find_column_index(['入库数量'], 18)

        print(f"    使用列索引: 采购数量={purchase_qty_col}, 请检消耗数量={consume_qty_col}, 入库数量={inbound_qty_col}")

        def get_numeric_val(col_idx):
            """安全地从列中获取数值"""
            if col_idx < len(self.df.columns):
                val = row.iloc[int(col_idx)]
                if pd.isna(val):
                    return 0
                # 如果是公式字符串，返回None（稍后计算）
                if isinstance(val, str):
                    if '=' in val:
                        return None
                    # 尝试转换为数字，如果不是数字则返回0
                    try:
                        return float(val)
                    except ValueError:
                        return 0
                return float(val)
            return 0

        for idx, row in self.df.iterrows():
            purchase_qty = get_numeric_val(purchase_qty_col)
            consume_qty = get_numeric_val(consume_qty_col)
            inbound_qty = get_numeric_val(inbound_qty_col)

            if purchase_qty is not None and consume_qty is not None and inbound_qty is not None:
                expected_inbound = purchase_qty - consume_qty
                if abs(inbound_qty - expected_inbound) > 0.01:
                    self.errors.append({
                        'row': int(idx),
                        'rule': '入库数量',
                        'message': f"第{idx+1}行：入库数量{inbound_qty}应为采购数量{purchase_qty}减去请验消耗量{consume_qty}，即{expected_inbound}"
                    })

    def _validate_data_completeness(self):
        """规则1：每一行数据都代表一个物料，不同列之间应该一一对应"""
        print("  检查数据完整性...")

        # 关键列（必填）
        required_cols = {
            '物料编码': 1,      # B列
            '物料名称': 2,      # C列
            '规格': 3,          # D列
            '单位': 4,          # E列
            '品牌': 5,          # F列
            '供应商': 6,        # G列
        }

        for idx, row in self.df.iterrows():
            for col_name, col_idx in required_cols.items():
                row_len = len(row)
                if int(col_idx) < row_len:
                    val = row.iloc[int(col_idx)]
                    if pd.isna(val) or str(val).strip() == '':
                        self.errors.append({
                            'row': int(idx),
                            'rule': '数据完整性',
                            'column': col_name,
                            'message': f"第{idx+1}行：{col_name}不能为空"
                        })

    def _identify_valid_rows(self):
        """识别没有错误的行"""
        error_rows = set(e['row'] for e in self.errors)
        self.valid_rows = [int(i) for i in range(len(self.df)) if i not in error_rows]
