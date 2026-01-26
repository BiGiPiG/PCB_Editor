from PyQt5.QtWidgets import QFileDialog
from domain.parameters import DrillingParams, TracksTrajectoryParams, BorderTrajectoryParams, MillingParams
from typing import List, Dict, Any


class ProjectManager:
    def __init__(self, kompas_service, settings_storage):
        self.kompas = kompas_service
        self.settings = settings_storage

    def generate_drilling_program(self, macro_id: str, params: DrillingParams) -> str:
        zero_macro = self.kompas.find_macro_by_type("Ноль станка")
        if not zero_macro:
            raise RuntimeError("Не найдена начальная точка ('Ноль станка')")
        return self.kompas.create_drilling_program(zero_macro, macro_id, params.depth, params.overrun, params.feedrate)

    def save_program_to_file(self, program: str, parent_window) -> bool:
        path, _ = QFileDialog.getSaveFileName(
            parent_window,
            "Сохранить файл сверловки",
            "",
            "CNC program (*.nc)"
        )
        if path:
            with open(path, "w") as f:
                f.write(program)
            return True
        return False

    def create_tracks_trajectory(self, macro_id: str, params: TracksTrajectoryParams):
        """Создаёт траекторию дорожек в Kompas"""
        zero_macro = self.kompas.find_macro_by_type("Ноль станка")
        if not zero_macro:
            raise RuntimeError("Не найдена начальная точка ('Ноль станка')")
        self.kompas.create_tracks_trajectory(
            zero_macro, macro_id,
            params.tool_diameter,
            params.line_count,
            params.overlap_percent
        )

    def create_border_trajectory(self, macro_id: str, params: BorderTrajectoryParams):
        """Создаёт траекторию границ в Kompas"""
        zero_macro = self.kompas.find_macro_by_type("Ноль станка")
        if not zero_macro:
            raise RuntimeError("Не найдена начальная точка ('Ноль станка')")
        self.kompas.create_border_trajectory(
            zero_macro, macro_id,
            params.tool_diameter,
            params.offset,
            params.contour_type.value
        )

    def generate_milling_program(self, macro_id: str, params: MillingParams) -> str:
        """Генерирует УП для фрезеровки"""
        zero_macro = self.kompas.find_macro_by_type("Ноль станка")
        if not zero_macro:
            raise RuntimeError("Не найдена начальная точка ('Ноль станка')")
        return self.kompas.create_milling_program(
            zero_macro, macro_id,
            params.depth, params.overrun, params.feedrate
        )

    def get_project_macros(self) -> List[Dict[str, Any]]:
        """
        Получает список макросов проекта в формате для TreeViewBuilder.
        """
        macros = []
        for macro in self.kompas.get_macros():
            try:
                keeper = self.kompas.kompas_api7_module.IPropertyKeeper(macro)
                name_prop = self.kompas.property_mng.GetProperty(self.kompas.doc, "Наименование")
                type_prop = self.kompas.property_mng.GetProperty(self.kompas.doc, "Тип")

                name_value = keeper.GetPropertyValue(name_prop, None, True, True)
                type_value = keeper.GetPropertyValue(type_prop, None, True, True)

                macros.append({
                    "name": name_value[1],
                    "type": type_value[1],
                    "id": macro
                })
            except Exception as e:
                print(f"Ошибка при получении свойств макроса: {e}")
                continue
        return macros