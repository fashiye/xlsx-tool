"""
DiffHighlighter: 文本差异高亮工具
提供：
- highlight_text_diff(str1, str2): 基于字符级 SequenceMatcher，返回带 HTML 高亮的左右字符串
- unified_diff_html(lines1, lines2): 返回统一 diff 格式的 HTML（带 <pre>）
- side_by_side_html(lines1, lines2): 生成并排行级 HTML，行差异带背景色
"""
import difflib
import html

class DiffHighlighter:
    def __init__(self):
        self.colors = {
            'added': '#90EE90',      # 浅绿色
            'removed': '#FFB6C1',    # 浅红色
            'changed': '#FFFFE0',    # 浅黄色
            'moved': '#ADD8E6'       # 浅蓝色
        }

    def highlight_text_diff(self, str1, str2):
        """
        字符级别高亮差异，返回 (html_left, html_right)
        使用 difflib.SequenceMatcher.get_opcodes()
        """
        if str1 is None:
            str1 = ""
        if str2 is None:
            str2 = ""
        matcher = difflib.SequenceMatcher(None, str1, str2)
        ops = matcher.get_opcodes()

        highlighted1 = []
        highlighted2 = []

        for tag, i1, i2, j1, j2 in ops:
            if tag == 'equal':
                highlighted1.append(html.escape(str1[i1:i2]))
                highlighted2.append(html.escape(str2[j1:j2]))
            elif tag == 'replace':
                highlighted1.append(f'<span style="background-color:{self.colors["changed"]}">{html.escape(str1[i1:i2])}</span>')
                highlighted2.append(f'<span style="background-color:{self.colors["changed"]}">{html.escape(str2[j1:j2])}</span>')
            elif tag == 'delete':
                highlighted1.append(f'<span style="background-color:{self.colors["removed"]}">{html.escape(str1[i1:i2])}</span>')
            elif tag == 'insert':
                highlighted2.append(f'<span style="background-color:{self.colors["added"]}">{html.escape(str2[j1:j2])}</span>')

        return ''.join(highlighted1), ''.join(highlighted2)

    def unified_diff_html(self, lines1, lines2):
        """
        生成 unified diff 的 HTML（以 <pre> 包裹）
        """
        diff = list(difflib.unified_diff(lines1, lines2, lineterm=''))
        # HTML escape and color added/removed lines
        out_lines = []
        for line in diff:
            esc = html.escape(line)
            if line.startswith('+') and not line.startswith('+++'):
                out_lines.append(f'<div style="background-color:{self.colors["added"]}">{esc}</div>')
            elif line.startswith('-') and not line.startswith('---'):
                out_lines.append(f'<div style="background-color:{self.colors["removed"]}">{esc}</div>')
            elif line.startswith('@@'):
                out_lines.append(f'<div style="background-color:{self.colors["changed"]}">{esc}</div>')
            else:
                out_lines.append(f'<div>{esc}</div>')
        return '<div style="font-family:monospace;">' + ''.join(out_lines) + '</div>'

    def side_by_side_html(self, lines1, lines2):
        """
        并排显示两组行，返回 HTML（左列为 lines1，右列为 lines2）
        使用 SequenceMatcher 的 opcodes 来标记行差异颜色
        """
        matcher = difflib.SequenceMatcher(None, lines1, lines2)
        left_html = []
        right_html = []
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                for line in lines1[i1:i2]:
                    left_html.append(f'<div style="background-color:#FFFFFF">{html.escape(line)}</div>')
                for line in lines2[j1:j2]:
                    right_html.append(f'<div style="background-color:#FFFFFF">{html.escape(line)}</div>')
            elif tag == 'replace':
                for line in lines1[i1:i2]:
                    left_html.append(f'<div style="background-color:{self.colors["removed"]}">{html.escape(line)}</div>')
                for line in lines2[j1:j2]:
                    right_html.append(f'<div style="background-color:{self.colors["added"]}">{html.escape(line)}</div>')
            elif tag == 'delete':
                for line in lines1[i1:i2]:
                    left_html.append(f'<div style="background-color:{self.colors["removed"]}">{html.escape(line)}</div>')
            elif tag == 'insert':
                for line in lines2[j1:j2]:
                    right_html.append(f'<div style="background-color:{self.colors["added"]}">{html.escape(line)}</div>')

        # Build a two-column table-like HTML using float layout
        left_col = '<div style="width:49%;float:left;border-right:1px solid #ddd;padding-right:6px;font-family:monospace;">' + ''.join(left_html) + '</div>'
        right_col = '<div style="width:49%;float:right;padding-left:6px;font-family:monospace;">' + ''.join(right_html) + '</div>'
        clear = '<div style="clear:both;"></div>'
        return left_col + right_col + clear