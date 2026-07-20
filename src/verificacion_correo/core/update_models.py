from enum import Enum
from dataclasses import dataclass, field
from typing import Optional


class UpdateStatus(Enum):
    SIN_ACTUALIZACION = "sin_actualizacion"
    ACTUALIZADO = "actualizado"
    SIN_INTERNET = "sin_internet"
    GIT_NO_DISPONIBLE = "git_no_disponible"
    REPOSITORIO_INVALIDO = "repositorio_invalido"
    CAMBIOS_LOCALES = "cambios_locales"
    HISTORIAL_DIVERGENTE = "historial_divergente"
    ERROR = "error"
    BLOQUEADO = "bloqueado"


@dataclass
class UpdateResult:
    status: UpdateStatus
    message: str
    previous_commit: Optional[str] = None
    current_commit: Optional[str] = None
    commits_updated: int = 0
    dependencies_updated: bool = False
    error_detail: Optional[str] = None


ERRORES_NO_BLOQUEANTES = {
    UpdateStatus.SIN_ACTUALIZACION,
    UpdateStatus.SIN_INTERNET,
    UpdateStatus.GIT_NO_DISPONIBLE,
    UpdateStatus.REPOSITORIO_INVALIDO,
}
