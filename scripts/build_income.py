"""Retired entry point: the previous builder depended on an excluded source."""

MESSAGE = "旧收入构建器已停用，未读取旧来源或匹配规则。请双击打开月度工作台.command，或运行 .venv/bin/python -m monthly.cli --month YYYY-MM。"

def build(folder, rules):
    raise RuntimeError(MESSAGE)

if __name__ == "__main__":
    raise SystemExit(MESSAGE)
