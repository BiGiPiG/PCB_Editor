from PyQt5.QtCore import QSettings
from domain.parameters import DrillingParams, TracksTrajectoryParams, BorderTrajectoryParams, ContourType, MillingParams


class SettingsStorage:
    def __init__(self):
        self.settings = QSettings("company", "pcb_editor")

    def save_drilling_defaults(self, params: DrillingParams):
        self.settings.setValue("depth", params.depth)
        self.settings.setValue("overrun", params.overrun)
        self.settings.setValue("feedrate", params.feedrate)

    def load_drilling_defaults(self) -> DrillingParams:
        return DrillingParams(
            depth=float(self.settings.value("depth", 0.0)),
            overrun=float(self.settings.value("overrun", 0.0)),
            feedrate=float(self.settings.value("feedrate", 0.0))
        )


    def save_tracks_defaults(self, params: TracksTrajectoryParams):
        self.settings.setValue("tracks_tool_diam", params.tool_diameter)
        self.settings.setValue("tracks_line_count", params.line_count)
        self.settings.setValue("tracks_overlap", params.overlap_percent)

    def load_tracks_defaults(self) -> TracksTrajectoryParams:
        return TracksTrajectoryParams(
            tool_diameter=float(self.settings.value("tracks_tool_diam", 0.0)),
            line_count=int(self.settings.value("tracks_line_count", 0)),
            overlap_percent=float(self.settings.value("tracks_overlap", 0.0))
        )


    def save_border_defaults(self, params: BorderTrajectoryParams):
        self.settings.setValue("border_tool_diam", params.tool_diameter)
        self.settings.setValue("border_offset", params.offset)
        self.settings.setValue("border_contour", params.contour_type.value)

    def load_border_defaults(self) -> BorderTrajectoryParams:
        return BorderTrajectoryParams(
            tool_diameter=float(self.settings.value("border_tool_diam", 0.0)),
            offset=float(self.settings.value("border_offset", 0.0)),
            contour_type=ContourType(self.settings.value("border_contour", "Внешний"))
        )

    def save_milling_defaults(self, params: MillingParams):
        self.settings.setValue("milling_depth", params.depth)
        self.settings.setValue("milling_overrun", params.overrun)
        self.settings.setValue("milling_feedrate", params.feedrate)

    def load_milling_defaults(self) -> MillingParams:
        return MillingParams(
            depth=float(self.settings.value("milling_depth", 0.0)),
            overrun=float(self.settings.value("milling_overrun", 0.0)),
            feedrate=float(self.settings.value("milling_feedrate", 0.0))
        )