from dataclasses import dataclass
from enum import Enum


@dataclass
class DrillingParams:
    depth: float
    overrun: float
    feedrate: float

@dataclass
class BordersParams:
    depth: float
    overrun: float
    feedrate: float

@dataclass
class TracksTrajectoryParams:
    tool_diameter: float
    line_count: int
    overlap_percent: float

class ContourType(Enum):
    EXTERNAL = "Внешний"
    INTERNAL = "Внутренний"

@dataclass
class BorderTrajectoryParams:
    tool_diameter: float
    offset: float
    contour_type: ContourType

@dataclass
class MillingParams:
    depth: float
    overrun: float
    feedrate: float