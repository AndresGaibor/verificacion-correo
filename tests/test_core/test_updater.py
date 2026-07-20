import tempfile
import subprocess
import os
import sys
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from verificacion_correo.core.updater import (
    check_for_updates,
    apply_update,
    _is_git_available,
    _get_repo_root,
    _is_repo_clean,
    _get_current_commit,
    _get_remote_commit,
    _count_commits_between,
    _acquire_lock,
    _release_lock,
    _get_lock_pid,
    _is_process_running,
    _get_file_hash,
    UPDATE_BRANCH,
)
from verificacion_correo.core.update_models import UpdateStatus


@pytest.fixture
def temp_git_repo():
    with tempfile.TemporaryDirectory() as tmpdir:
        subprocess.run(["git", "init"], cwd=tmpdir, capture_output=True)
        subprocess.run(["git", "config", "user.email", "test@test.com"], cwd=tmpdir, capture_output=True)
        subprocess.run(["git", "config", "user.name", "Test"], cwd=tmpdir, capture_output=True)
        Path(tmpdir, "test.txt").write_text("hello")
        subprocess.run(["git", "add", "."], cwd=tmpdir, capture_output=True)
        subprocess.run(["git", "commit", "-m", "init"], cwd=tmpdir, capture_output=True)
        yield Path(tmpdir)


class TestCheckForUpdates:
    def test_check_updates_no_remote(self, temp_git_repo):
        result = check_for_updates(temp_git_repo)
        assert result.status == UpdateStatus.SIN_INTERNET

    def test_check_updates_git_not_available(self):
        with patch("verificacion_correo.core.updater._is_git_available", return_value=False):
            result = check_for_updates(Path("/some/path"))
            assert result.status == UpdateStatus.GIT_NO_DISPONIBLE

    def test_check_updates_not_a_git_repo(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            result = check_for_updates(Path(tmpdir))
            assert result.status == UpdateStatus.REPOSITORIO_INVALIDO


class TestIsGitAvailable:
    def test_git_available_returns_bool(self):
        result = _is_git_available()
        assert isinstance(result, bool)


class TestGetRepoRoot:
    def test_get_repo_root_valid(self, temp_git_repo):
        root = _get_repo_root(temp_git_repo)
        assert root is not None
        assert root.resolve() == temp_git_repo.resolve()

    def test_get_repo_root_invalid(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = _get_repo_root(Path(tmpdir))
            assert root is None


class TestIsRepoClean:
    def test_is_repo_clean_true(self, temp_git_repo):
        clean, reason = _is_repo_clean(temp_git_repo)
        assert clean is True
        assert reason == ""

    def test_is_repo_clean_false_with_changes(self, temp_git_repo):
        Path(temp_git_repo, "new_file.txt").write_text("new content")
        clean, reason = _is_repo_clean(temp_git_repo)
        assert clean is False
        assert "Cambios locales" in reason


class TestGetCurrentCommit:
    def test_get_current_commit(self, temp_git_repo):
        commit = _get_current_commit(temp_git_repo)
        assert commit is not None
        assert len(commit) == 40

    def test_get_current_commit_invalid_repo(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            commit = _get_current_commit(Path(tmpdir))
            assert commit is None


class TestGetRemoteCommit:
    def test_get_remote_commit_no_remote(self, temp_git_repo):
        commit = _get_remote_commit(temp_git_repo)
        assert commit is None

    def test_get_remote_commit_with_fake_remote(self, temp_git_repo):
        subprocess.run(["git", "remote", "add", "origin", "https://github.com/andresgaibor/verificacion-correo.git"], cwd=temp_git_repo, capture_output=True)
        subprocess.run(["git", "fetch", "origin", "main"], cwd=temp_git_repo, capture_output=True)
        commit = _get_remote_commit(temp_git_repo, "origin/main")
        assert commit is None or isinstance(commit, str)


class TestCountCommitsBetween:
    def test_count_commits_same_commit(self, temp_git_repo):
        commit = _get_current_commit(temp_git_repo)
        count = _count_commits_between(temp_git_repo, commit, commit)
        assert count == 0


class TestAcquireReleaseLock:
    def test_acquire_lock_returns_bool(self):
        with patch("verificacion_correo.core.updater.get_lock_path") as mock_lock_path:
            with tempfile.TemporaryDirectory() as tmpdir:
                lock_file = Path(tmpdir) / "lock"
                mock_lock_path.return_value = lock_file
                result = _acquire_lock()
                assert isinstance(result, bool)

    def test_release_lock_no_error(self):
        with patch("verificacion_correo.core.updater.get_lock_path") as mock_lock_path:
            with tempfile.TemporaryDirectory() as tmpdir:
                lock_file = Path(tmpdir) / "lock"
                mock_lock_path.return_value = lock_file
                lock_file.write_text("12345")
                _release_lock()


class TestGetLockPid:
    def test_get_lock_pid_no_file(self):
        with patch("verificacion_correo.core.updater.get_lock_path") as mock_lock_path:
            with tempfile.TemporaryDirectory() as tmpdir:
                lock_file = Path(tmpdir) / "nonexistent.lock"
                mock_lock_path.return_value = lock_file
                pid = _get_lock_pid()
                assert pid is None

    def test_get_lock_pid_invalid_content(self):
        with patch("verificacion_correo.core.updater.get_lock_path") as mock_lock_path:
            with tempfile.TemporaryDirectory() as tmpdir:
                lock_file = Path(tmpdir) / "lock"
                mock_lock_path.return_value = lock_file
                lock_file.write_text("not_a_number")
                pid = _get_lock_pid()
                assert pid is None


class TestIsProcessRunning:
    def test_is_process_running_invalid_pid(self):
        result = _is_process_running(999999)
        assert result is False

    def test_is_process_running_current_pid(self):
        result = _is_process_running(os.getpid())
        assert result is True


class TestGetFileHash:
    def test_get_file_hash_nonexistent(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            hash_val = _get_file_hash(Path(tmpdir) / "nonexistent.txt")
            assert hash_val == ""

    def test_get_file_hash_valid_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            test_file = Path(tmpdir) / "test.txt"
            test_file.write_text("hello world")
            hash_val = _get_file_hash(test_file)
            assert hash_val != ""
            assert len(hash_val) == 32


class TestApplyUpdate:
    def test_apply_update_no_remote(self, temp_git_repo):
        with patch("verificacion_correo.core.updater._acquire_lock", return_value=True):
            with patch("verificacion_correo.core.updater._release_lock"):
                with patch("verificacion_correo.core.updater._get_current_commit", return_value="abc123"):
                    result = apply_update(temp_git_repo)
                    assert result.status == UpdateStatus.ERROR

    def test_apply_update_locked(self):
        with patch("verificacion_correo.core.updater._acquire_lock", return_value=False):
            with tempfile.TemporaryDirectory() as tmpdir:
                result = apply_update(Path(tmpdir))
                assert result.status == UpdateStatus.BLOQUEADO
