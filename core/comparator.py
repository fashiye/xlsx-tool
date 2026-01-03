"""
核心比较器：管理工作簿、选择单元格、直接比较、公式验证、导出结果（基本）
"""
from core.excel_reader import load_workbook_all_sheets
from core.string_comparator import StringComparator
from core.validator import validate_formula
import pandas as pd
import re

class ExcelComparator:
    def __init__(self):
        self.workbooks = {}  # alias -> { 'path':..., 'sheets': {name:DataFrame} }
        self.string_comparator = StringComparator()

    def load_workbook(self, filepath, alias=None):
        alias = alias or filepath
        sheets = load_workbook_all_sheets(filepath)
        self.workbooks[alias] = {
            'path': filepath,
            'sheets': sheets
        }

    def list_sheets(self, alias):
        if alias not in self.workbooks:
            return []
        return list(self.workbooks[alias]['sheets'].keys())

    def get_sheet_dataframe(self, alias, sheet_name):
        if alias not in self.workbooks:
            raise ValueError(f"未加载工作簿: {alias}")
        sheets = self.workbooks[alias]['sheets']
        if sheet_name not in sheets:
            raise ValueError(f"工作表 {sheet_name} 不存在")
        df = sheets[sheet_name].copy()
        # reset index to simple 0..n-1 to align with model
        df = df.reset_index(drop=True)
        return df

    cell_ref_re = re.compile(r'^([A-Za-z]+)(\d+)$')

    @staticmethod
    def col_letters_to_index(col_letters):
        """A -> 0, B -> 1, AA -> 26"""
        col_letters = col_letters.upper()
        idx = 0
        for ch in col_letters:
            idx = idx * 26 + (ord(ch) - ord('A') + 1)
        return idx - 1

    def parse_range(self, rng):
        """
        支持单元格或矩形:
        - "A1" -> (col0,row0, col0,row0)
        - "A1:C10" -> (col0,row0, col2,row9)
        1-based rows in Excel -> we convert to 0-based
        """
        parts = rng.split(':')
        if len(parts) == 1:
            m = self.cell_ref_re.match(parts[0])
            if not m:
                raise ValueError("无效单元格地址")
            c = self.col_letters_to_index(m.group(1))
            r = int(m.group(2)) - 1
            return c, r, c, r
        elif len(parts) == 2:
            m1 = self.cell_ref_re.match(parts[0])
            m2 = self.cell_ref_re.match(parts[1])
            if not m1 or not m2:
                raise ValueError("无效单元格范围")
            c1 = self.col_letters_to_index(m1.group(1))
            r1 = int(m1.group(2)) - 1
            c2 = self.col_letters_to_index(m2.group(1))
            r2 = int(m2.group(2)) - 1
            # normalize
            return min(c1,c2), min(r1,r2), max(c1,c2), max(r1,r2)
        else:
            raise ValueError("不支持的范围格式")

    def select_cells(self, workbook_alias, sheet_name, rng):
        """
        返回 pandas.DataFrame 对应范围（如果超出 sheet 大小，返回可用交集）
        """
        df = self.get_sheet_dataframe(workbook_alias, sheet_name)
        c1, r1, c2, r2 = self.parse_range(rng)
        # pandas uses columns as names; we convert by position
        max_cols = df.shape[1]
        max_rows = df.shape[0]
        # if sheet has no columns, return empty df with 0 cols
        if max_cols == 0:
            return pd.DataFrame()
        c1 = max(0, c1); c2 = min(max_cols - 1, c2)
        r1 = max(0, r1); r2 = min(max_rows - 1, r2)
        cols = list(df.columns[c1:c2+1])
        sub = df.loc[r1:r2, cols].reset_index(drop=True)
        # ensure column names are present
        return sub

    def compare_direct(self, df1, df2, options=None):
        """
        逐单元格比较两个 DataFrame
        返回: result_df (以 df1 的 shape 为基准, 或最大行列的表), result_map {(r,c): 'equal'/'diff'/...}
        选项:
          tolerance: 数值容差
          ignore_case: bool
        """
        options = options or {}
        tol = float(options.get('tolerance', 0))
        ignore_case = bool(options.get('ignore_case', False))

        # determine final table shape: use max of rows/cols to cover both
        rows = max(df1.shape[0], df2.shape[0])
        cols = max(df1.shape[1], df2.shape[1])

        # ensure columns for result
        col_names = []
        for i in range(cols):
            # try get column name from df1 else df2 else generic
            name = None
            if i < df1.shape[1]:
                name = str(df1.columns[i])
            elif i < df2.shape[1]:
                name = str(df2.columns[i])
            else:
                name = f"COL_{i}"
            col_names.append(name)

        # build result dataframe by picking df1 where available otherwise df2
        result_vals = []
        result_map = {}
        for r in range(rows):
            row_vals = []
            for c in range(cols):
                v1 = df1.iat[r, c] if (r < df1.shape[0] and c < df1.shape[1]) else None
                v2 = df2.iat[r, c] if (r < df2.shape[0] and c < df2.shape[1]) else None

                status = 'empty'
                display = v1 if v1 is not None else v2
                # both NaN handling
                import pandas as _pd
                if _pd.isna(v1) and _pd.isna(v2):
                    status = 'equal'
                else:
                    # numeric compare
                    if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):
                        diff = abs float(v1 - v2) if False else abs(v1 - v2)
                        if diff <= tol:
                            status = 'equal'
                        else:
                            status = 'diff'
                    else:
                        # convert to str and compare (respect ignore_case)
                        s1 = "" if v1 is None or (isinstance(v1, float) and _pd.isna(v1)) else str(v1)
                        s2 = "" if v2 is None or (isinstance(v2, float) and _pd.isna(v2)) else str(v2)
                        if ignore_case:
                            s1c = s1.lower()
                            s2c = s2.lower()
                        else:
                            s1c = s1; s2c = s2
                        if s1c == s2c:
                            status = 'equal'
                        else:
                            status = 'diff'
                row_vals.append(display)
                result_map[(r, c)] = status
            result_vals.append(row_vals)

        result_df = pd.DataFrame(result_vals, columns=col_names)
        return result_df, result_map

    def validate_formula(self, cells_dict, formula, expected_value, options=None):
        options = options or {}
        tol = float(options.get('tolerance', 0.0))
        return validate_formula(cells_dict, formula, expected_value, tolerance=tol)

    def export_results(self, result_df, output_path, format='excel'):
        if format == 'excel':
            result_df.to_excel(output_path, index=False)
        elif format == 'csv':
            result_df.to_csv(output_path, index=False)
        else:
            raise ValueError("不支持的导出格式")