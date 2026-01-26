from PyQt5.QtWidgets import QTreeWidget, QTreeWidgetItem
from typing import List, Dict, Any


class TreeViewBuilder:
    @staticmethod
    def build_project_tree(
            tree_widget: QTreeWidget,
            project_name: str,
            macros: List[Dict[str, Any]]
    ) -> None:
        """
        Строит дерево проекта в указанном QTreeWidget.

        :param tree_widget: Виджет дерева для заполнения
        :param project_name: Название проекта (корневой элемент)
        :param macros: Список макросов в формате:
            [
                {"name": "Ноль станка", "type": "Ноль станка", "id": macro_object},
                {"name": "Слой1", "type": "Дорожки", "id": macro_object},
                ...
            ]
        """
        tree_widget.clear()

        root = QTreeWidgetItem(tree_widget)
        root.setText(0, project_name)
        root.setExpanded(True)

        for macro in macros:
            display_name = f"{macro['name']} ({macro['type']})"
            item = QTreeWidgetItem(root)
            item.setText(0, display_name)
            item.setData(1, 0, macro["id"])