import pytest
from verificacion_correo.core.update_models import UpdateStatus, UpdateResult, ERRORES_NO_BLOQUEANTES


def test_update_status_enum_values():
    assert UpdateStatus.SIN_ACTUALIZACION.value == "sin_actualizacion"
    assert UpdateStatus.ACTUALIZADO.value == "actualizado"
    assert UpdateStatus.SIN_INTERNET.value == "sin_internet"


def test_update_result_dataclass():
    result = UpdateResult(
        status=UpdateStatus.SIN_ACTUALIZACION,
        message="No hay actualizaciones",
        previous_commit="abc123",
        current_commit="abc123",
        commits_updated=0,
    )
    assert result.status == UpdateStatus.SIN_ACTUALIZACION
    assert result.commits_updated == 0


def test_update_result_defaults():
    result = UpdateResult(status=UpdateStatus.ACTUALIZADO, message="OK")
    assert result.previous_commit is None
    assert result.current_commit is None
    assert result.commits_updated == 0
    assert result.dependencies_updated is False
    assert result.error_detail is None


def test_update_status_git_no_disponible():
    assert UpdateStatus.GIT_NO_DISPONIBLE.value == "git_no_disponible"


def test_update_status_repositorio_invalido():
    assert UpdateStatus.REPOSITORIO_INVALIDO.value == "repositorio_invalido"


def test_update_status_cambios_locales():
    assert UpdateStatus.CAMBIOS_LOCALES.value == "cambios_locales"


def test_update_status_historial_divergente():
    assert UpdateStatus.HISTORIAL_DIVERGENTE.value == "historial_divergente"


def test_update_status_error():
    assert UpdateStatus.ERROR.value == "error"


def test_update_status_bloqueado():
    assert UpdateStatus.BLOQUEADO.value == "bloqueado"


def test_errores_no_bloqueantes():
    assert UpdateStatus.SIN_ACTUALIZACION in ERRORES_NO_BLOQUEANTES
    assert UpdateStatus.SIN_INTERNET in ERRORES_NO_BLOQUEANTES
    assert UpdateStatus.GIT_NO_DISPONIBLE in ERRORES_NO_BLOQUEANTES
    assert UpdateStatus.REPOSITORIO_INVALIDO in ERRORES_NO_BLOQUEANTES
    assert UpdateStatus.ACTUALIZADO not in ERRORES_NO_BLOQUEANTES
    assert UpdateStatus.ERROR not in ERRORES_NO_BLOQUEANTES
