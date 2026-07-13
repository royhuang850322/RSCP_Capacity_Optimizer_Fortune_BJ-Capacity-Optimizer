"""Disabled legacy CLI for Fortune BJ input template generation."""
from __future__ import annotations


def main() -> None:
    raise SystemExit(
        "生成/刷新数据导入模板功能已停用。当前 数据导入 文件夹就是正式数据源，"
        "工具运行时只读取输入数据，不再从 Project_Stuff 抽取或覆盖源文件。"
    )


if __name__ == "__main__":
    main()
