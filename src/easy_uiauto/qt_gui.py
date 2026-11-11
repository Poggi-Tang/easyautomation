"""Qt knowledge GUI entry; legacy Tk classes remain import-compatible."""
from __future__ import annotations


def configure_application(app):
    from PySide6.QtGui import QFont
    app.setFont(QFont("Microsoft YaHei UI",10))
    app.setStyleSheet("QPushButton,QToolButton { min-height: 26px; padding: 0 9px; } QLineEdit,QComboBox { min-height: 24px; } QTreeView { outline: none; } QTreeView::item { min-height: 25px; } QHeaderView::section { padding: 4px; }")


def main():
    import argparse
    import sys
    from pathlib import Path
    from PySide6.QtWidgets import QApplication
    from .qt_ui.window import KnowledgeWindow

    parser=argparse.ArgumentParser(description="选择软件窗口、学习区域、编排组合动作")
    parser.add_argument("--vault",type=Path)
    args=parser.parse_args()
    app=QApplication.instance() or QApplication(sys.argv[:1])
    configure_application(app)
    window=KnowledgeWindow(args.vault);window.show()
    raise SystemExit(app.exec())


if __name__=="__main__":main()
