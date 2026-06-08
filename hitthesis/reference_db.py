"""参考文献数据库，支持 BibTeX 解析与 GB/T 7714-2015 格式化"""

import re
import bibtexparser


class ReferenceDB:
    def __init__(self, bib_path: str):
        """解析 .bib 文件，加载所有条目"""
        self.entries = {}  # key -> entry dict
        self._citation_order = []  # 按引用顺序记录 key（可重复）
        self._citation_index = {}  # key -> first occurrence order (1-based)

        with open(bib_path, encoding="utf-8") as f:
            bib_db = bibtexparser.load(f)

        for entry in bib_db.entries:
            key = entry.get("ID", entry.get("key"))
            self.entries[key] = dict(entry)

    def cite(self, key: str) -> int:
        """记录一次引用，返回序号（1-based），重复引用返回已有序号"""
        if key not in self._citation_index:
            self._citation_index[key] = len(self._citation_index) + 1
        self._citation_order.append(key)
        return self._citation_index[key]

    def get_cited_keys(self):
        """返回按引用顺序去重的 key 列表"""
        seen = set()
        result = []
        for k in self._citation_order:
            if k not in seen:
                seen.add(k)
                result.append(k)
        return result

    def get_references_in_citation_order(self):
        """返回按引用顺序格式化的文献字符串列表（不含编号）"""
        return [self.format_entry(key) for key in self.get_cited_keys()]

    def format_entry(self, key: str) -> str:
        """将条目格式化为 GB/T 7714-2015 字符串"""
        if key not in self.entries:
            return ""
        entry = self.entries[key]
        entry_type = entry.get("ENTRYTYPE", "").lower()
        if entry_type == "article":
            return self._format_article(entry)
        elif entry_type == "book":
            return self._format_book(entry)
        else:
            # 默认当作文章处理
            return self._format_article(entry)

    def _format_article(self, entry) -> str:
        """格式化期刊文章 [J]

        格式：作者. 题名[J]. 刊名, 年, 卷(期): 页码.
        """
        authors = self._format_author(entry.get("author", ""))
        title = entry.get("title", "")
        journal = entry.get("journal", "")
        year = entry.get("year", "")
        volume = entry.get("volume", "")
        number = entry.get("number", "")
        pages = entry.get("pages", "")

        # 判断是否中文文献
        is_cn = self._is_chinese(title) or self._is_chinese(journal)
        sep = "，" if is_cn else ", "
        colon = "：" if is_cn else ": "

        # 构建卷号(期号)
        vol_part = volume
        if number:
            vol_part += f"({number})"

        # 组装各部分
        parts = []
        if authors:
            parts.append(authors + ".")
        if title:
            parts.append(title + "[J].")
        # 刊名, 年, 卷(期): 页码 — 用逗号连接中间部分，最后加句点
        mid = []
        if journal:
            mid.append(journal)
        if year:
            mid.append(year)
        if vol_part:
            vol_str = vol_part
            if pages:
                vol_str += f"{colon}{pages}"
            mid.append(vol_str)
        elif pages:
            mid.append(pages)
        if mid:
            parts.append(sep.join(mid) + ".")

        result = "".join(parts)
        if not result.endswith("."):
            result += "."
        return result

    def _format_book(self, entry) -> str:
        """格式化图书 [M]

        格式：作者. 题名[M]. 出版地: 出版社, 出版年.
        """
        authors = self._format_author(entry.get("author", ""))
        title = entry.get("title", "")
        address = entry.get("address", "")
        publisher = entry.get("publisher", "")
        year = entry.get("year", "")

        # 判断是否中文文献
        is_cn = self._is_chinese(title) or self._is_chinese(publisher)
        colon = "：" if is_cn else ": "
        sep = "，" if is_cn else ", "

        parts = []
        if authors:
            parts.append(authors + ".")
        if title:
            parts.append(title + "[M].")
        if address and publisher:
            parts.append(f"{address}{colon}{publisher}{sep}")
        elif publisher:
            parts.append(f"{publisher}{sep}")
        if year:
            if is_cn:
                parts.append(f"{year}.")
            else:
                parts.append(f" {year}.")

        result = "".join(parts)
        if not result.endswith("."):
            result += "."
        return result

    def _format_author(self, author_str: str) -> str:
        """格式化作者字符串

        - 外文：姓全大写，名取首字母缩写，空格分隔
          例如：Gross, W. A. -> GROSS W A
        - 中文汉字：直录，多作者用 `，`（全角逗号）分隔
        - 超过3个：中文 `， 等`（无句点），外文 `, et al`（无句点）
        """
        # 分割多位作者
        authors = re.split(r"\s+and\s+", author_str.strip())
        authors = [a.strip() for a in authors if a.strip()]

        formatted = []
        has_chinese = False
        has_foreign = False

        for author in authors:
            if self._is_chinese(author):
                # 中文作者直录
                formatted.append(author)
                has_chinese = True
            else:
                # 外文作者：Last, F. M. -> LAST F M
                parts = author.split(",")
                if len(parts) >= 2:
                    last = parts[0].strip().upper()
                    initials_raw = parts[1].strip()
                    # 提取首字母，移除所有非字母字符
                    initials = re.sub(r"[^a-zA-Z]", "", initials_raw)
                    # 转为大写，空格分隔
                    initials_formatted = " ".join(c.upper() for c in initials)
                    formatted.append(f"{last} {initials_formatted}")
                else:
                    formatted.append(author.upper())
                has_foreign = True

        # 选择分隔符
        if has_chinese and not has_foreign:
            sep = "，"
        else:
            sep = ", "

        # 超过3个时省略
        if len(formatted) > 3:
            if has_chinese and not has_foreign:
                return sep.join(formatted[:3]) + "，等"
            else:
                return sep.join(formatted[:3]) + ", et al"
        elif len(formatted) == 3:
            return sep.join(formatted)
        elif len(formatted) == 2:
            return sep.join(formatted)
        elif len(formatted) == 1:
            return formatted[0]
        else:
            return ""

    def _is_chinese(self, text: str) -> bool:
        """判断是否全为中文汉字（不含拉丁字母）"""
        # 包含汉字且不含拉丁字母 A-Z a-z
        has_chinese = bool(re.search(r"[一-鿿]", text))
        has_latin = bool(re.search(r"[A-Za-z]", text))
        return has_chinese and not has_latin
