from __future__ import annotations

import os
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
    combo.setEditable(os.environ.get("EASY_UIAUTO_QT_EDITABLE_COMBO") == "1")
    status = QLabel("idle")
    status.setObjectName("status_label")
    status.setAccessibleName("idle")

    combo_status = QLabel("combo:Alpha")
    combo_status.setObjectName("combo_status")
    combo_status.setAccessibleName("combo:Alpha")
    tree_status = QLabel("tree:collapsed")
    tree_status.setObjectName("tree_status")
    tree_status.setAccessibleName("tree:collapsed")
    tree_selection_status = QLabel("tree-selection:none")
    tree_selection_status.setObjectName("tree_selection_status")
    tree_selection_status.setAccessibleName("tree-selection:none")
    table_status = QLabel("table:none")
    table_status.setObjectName("table_status")
    table_status.setAccessibleName("table:none")
    tab_status = QLabel("tab:First")
    tab_status.setObjectName("tab_status")
    tab_status.setAccessibleName("tab:First")

    tree = QTreeWidget()
    tree.setObjectName("fixture_tree")
    tree.setHeaderLabel("Items")
    root = QTreeWidgetItem(["Root"])
    root.addChild(QTreeWidgetItem(["Child"]))
    tree.addTopLevelItem(root)
    root.setExpanded(os.environ.get("EASY_UIAUTO_QT_TREE_EXPANDED") == "1")
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

    def set_status(label, value):
        label.setText(value)
        label.setAccessibleName(value)

    def update_tree_selection():
        selected = tree.selectedItems()
        value = selected[0].text(0) if selected else "none"
        set_status(tree_selection_status, f"tree-selection:{value}")

    def update_table_selection():
        selected = table.selectedItems()
        value = selected[0].text() if selected else "none"
        set_status(table_status, f"table:{value}")

    combo.currentTextChanged.connect(
        lambda value: set_status(combo_status, f"combo:{value}")
    )
    tree.itemExpanded.connect(lambda _item: set_status(tree_status, "tree:expanded"))
    tree.itemCollapsed.connect(lambda _item: set_status(tree_status, "tree:collapsed"))
    tree.itemSelectionChanged.connect(update_tree_selection)
    table.itemSelectionChanged.connect(update_table_selection)
    tabs.currentChanged.connect(
        lambda index: set_status(tab_status, f"tab:{tabs.tabText(index)}")
    )
    set_status(tree_status, "tree:expanded" if root.isExpanded() else "tree:collapsed")

    button.clicked.connect(apply_value)
    layout.addWidget(edit, 0, 0, 1, 2)
    layout.addWidget(button, 1, 0)
    layout.addWidget(checkbox, 1, 1)
    layout.addWidget(combo, 2, 0, 1, 2)
    layout.addWidget(status, 3, 0, 1, 2)
    layout.addWidget(combo_status, 4, 0)
    layout.addWidget(tree_status, 4, 1)
    layout.addWidget(tree_selection_status, 5, 0)
    layout.addWidget(table_status, 5, 1)
    layout.addWidget(tab_status, 6, 0, 1, 2)
    layout.addWidget(tabs, 7, 0, 1, 2)
    layout.addWidget(tree, 8, 0)
    layout.addWidget(table, 8, 1)
    window.resize(760, 700)
    window.show()
    print("READY", flush=True)
    raise SystemExit(app.exec())


if __name__ == "__main__":
    main()
