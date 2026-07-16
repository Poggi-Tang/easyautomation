from __future__ import annotations

import sys

from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QWidget,
)


def main():
    app = QApplication(sys.argv)
    title = sys.argv[1] if len(sys.argv) > 1 else "Easy UIAuto Qt Fixture"
    window = QWidget()
    window.setWindowTitle(title)
    window.setObjectName("fixture_window")
    layout = QGridLayout(window)

    edit = QLineEdit()
    edit.setObjectName("input_edit")
    edit.setAccessibleName("Input")
    button = QPushButton("Apply")
    button.setObjectName("apply_button")
    checkbox = QCheckBox("Enabled")
    checkbox.setObjectName("enabled_checkbox")
    combo = QComboBox()
    combo.setObjectName("mode_combo")
    combo.addItems(["Alpha", "Beta"])
    status = QLabel("idle")
    status.setObjectName("status_label")
    status.setAccessibleName("idle")

    tree = QTreeWidget()
    tree.setObjectName("fixture_tree")
    tree.setHeaderLabel("Items")
    root = QTreeWidgetItem(["Root"])
    root.addChild(QTreeWidgetItem(["Child"]))
    tree.addTopLevelItem(root)
    table = QTableWidget(2, 2)
    table.setObjectName("fixture_table")
    table.setHorizontalHeaderLabels(["Name", "Value"])
    table.setItem(0, 0, QTableWidgetItem("row-a"))
    table.setItem(0, 1, QTableWidgetItem("1"))
    tabs = QTabWidget()
    tabs.setObjectName("fixture_tabs")
    tabs.addTab(QLabel("First page"), "First")
    tabs.addTab(QLabel("Second page"), "Second")

    def apply_value():
        value = f"applied:{edit.text()}"
        status.setText(value)
        status.setAccessibleName(value)

    button.clicked.connect(apply_value)
    layout.addWidget(edit, 0, 0, 1, 2)
    layout.addWidget(button, 1, 0)
    layout.addWidget(checkbox, 1, 1)
    layout.addWidget(combo, 2, 0, 1, 2)
    layout.addWidget(status, 3, 0, 1, 2)
    layout.addWidget(tabs, 4, 0, 1, 2)
    layout.addWidget(tree, 5, 0)
    layout.addWidget(table, 5, 1)
    window.resize(760, 620)
    window.show()
    print("READY", flush=True)
    raise SystemExit(app.exec())


if __name__ == "__main__":
    main()
