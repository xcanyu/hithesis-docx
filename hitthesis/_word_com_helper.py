"""Word COM 子进程 helper：在独立进程中执行 update_fields，避免崩溃影响主进程。

为什么独立进程：
  Word COM 在某些环境（pytest、特殊用户配置、残留 Word 进程）会触发
  0x800706be 等 fatal exception，无法被 Python except 捕获。
  用 subprocess 隔离后，崩溃只影响子进程，主测试/主流程继续。
"""
import os
import sys
import time


def main():
    if len(sys.argv) < 2:
        print("Usage: _word_com_helper.py <docx_path>", file=sys.stderr)
        sys.exit(1)
    abs_path = os.path.abspath(sys.argv[1])
    if not os.path.exists(abs_path):
        print(f"File not found: {abs_path}", file=sys.stderr)
        sys.exit(1)

    import win32com.client
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        doc = word.Documents.Open(abs_path)
        doc.Fields.Update()
        for toc in doc.TablesOfContents:
            toc.Update()
        doc.Repaginate()
        doc.Fields.Update()
        doc.Save()
        time.sleep(0.3)
        doc.Close(False)
    finally:
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
