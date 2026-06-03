from dataclasses import dataclass, asdict

@dataclass
class ScriptModel:
    name: str
    path: str
    description: str = ""
    excel_path: str = ""
    tab: str = "Schede"
    group: str = ""
    notes: str = ""
    last_executed: str = "Mai"
    order: int = 0

    @classmethod
    def from_dict(cls, data: dict):
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

    def to_dict(self):
        return asdict(self)
