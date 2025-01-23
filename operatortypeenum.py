from enum import Enum, auto

class OperatorType(Enum):
    FOLDER = auto()
    FILE = auto()
    ALL = auto()