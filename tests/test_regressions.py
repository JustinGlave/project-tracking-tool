from __future__ import annotations

import shutil
import unittest
import uuid
import json
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

from project_tracker_backend import (
    DEFAULT_TASKS,
    ProjectRecord,
    ProjectTrackerBackend,
    parse_currency,
)
from user_auth import AuthStoreError, UserManager
from updater import (
    UpdatePackageError,
    _build_update_powershell_script,
    _validate_update_zip,
)


class TempWorkspaceTest(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = Path(".test_tmp") / uuid.uuid4().hex
        self.tmp.mkdir(parents=True)
        self.addCleanup(lambda: shutil.rmtree(self.tmp, ignore_errors=True))


class AuthRegressionTests(TempWorkspaceTest):
    def test_corrupt_user_store_does_not_look_empty(self) -> None:
        users_path = self.tmp / "users.json"
        users_path.write_text("{bad json", encoding="utf-8")

        with self.assertRaises(AuthStoreError):
            UserManager(users_path)

    def test_first_created_user_is_admin(self) -> None:
        manager = UserManager(self.tmp / "users.json")
        manager.create_user("alice", "password123", role="user")

        user = manager.get_user("alice")
        assert user is not None
        self.assertEqual(user.role, "admin")

    def test_remember_me_requires_token_and_reset_invalidates_it(self) -> None:
        users_path = self.tmp / "users.json"
        manager = UserManager(users_path)
        manager.create_user("alice", "password123")

        token = manager.create_session_token("alice")
        user_data = json.loads(users_path.read_text(encoding="utf-8"))
        self.assertNotIn("session_token_hash", user_data["alice"])
        self.assertNotIn("session_token_expires_at", user_data["alice"])
        self.assertIsNotNone(manager.authenticate_session("alice", token))
        self.assertIsNone(manager.authenticate_session("alice", "not-the-token"))
        self.assertIsNone(manager.authenticate_session("alice", ""))

        manager.reset_password("alice", "temporary123")
        self.assertIsNone(manager.authenticate_session("alice", token))

    def test_remember_me_token_expires(self) -> None:
        users_path = self.tmp / "users.json"
        manager = UserManager(users_path)
        manager.create_user("alice", "password123")
        token = manager.create_session_token("alice")

        sessions_path = users_path.with_name("user_sessions.json")
        data = json.loads(sessions_path.read_text(encoding="utf-8"))
        data["alice"]["expires_at"] = (
            datetime.now() - timedelta(days=1)
        ).replace(microsecond=0).isoformat(sep=" ")
        sessions_path.write_text(json.dumps(data), encoding="utf-8")

        self.assertIsNone(manager.authenticate_session("alice", token))

    def test_v180_session_fields_are_migrated_out_of_users_file(self) -> None:
        users_path = self.tmp / "users.json"
        token = "saved-token"
        expires_at = (datetime.now() + timedelta(days=1)).replace(
            microsecond=0
        ).isoformat(sep=" ")
        users_path.write_text(
            json.dumps(
                {
                    "alice": {
                        "username": "alice",
                        "password_hash": "hash",
                        "salt": "00",
                        "must_change_password": False,
                        "created_at": "",
                        "role": "admin",
                        "session_token_hash": token,
                        "session_token_expires_at": expires_at,
                    }
                }
            ),
            encoding="utf-8",
        )

        UserManager(users_path)

        users_data = json.loads(users_path.read_text(encoding="utf-8"))
        self.assertNotIn("session_token_hash", users_data["alice"])
        self.assertNotIn("session_token_expires_at", users_data["alice"])
        sessions_data = json.loads(
            users_path.with_name("user_sessions.json").read_text(encoding="utf-8")
        )
        self.assertEqual(sessions_data["alice"]["token_hash"], token)


class BackendRegressionTests(TempWorkspaceTest):
    def test_parse_currency_handles_user_formatted_values(self) -> None:
        self.assertEqual(parse_currency("$1,234.56"), 1234.56)
        self.assertEqual(parse_currency("(1,234.56)"), -1234.56)
        self.assertEqual(parse_currency("not money"), 0.0)

    def test_excel_export_accepts_formatted_contract_value(self) -> None:
        backend = ProjectTrackerBackend(self.tmp / "data.json")
        project_id = backend.create_project(
            ProjectRecord(
                job_name="Formatted Contract",
                job_number="12345",
                contract_value="$1,234.56",
            )
        )

        output_path = backend.export_project_to_excel(project_id, self.tmp / "out.xlsx")
        self.assertTrue(output_path.exists())

    def test_dashboard_task_counts_exclude_test_jobs(self) -> None:
        backend = ProjectTrackerBackend(self.tmp / "data.json")
        backend.create_project(ProjectRecord(job_name="Real", job_number="R"))
        backend.create_project(ProjectRecord(job_name="Training", job_number="T", is_test=True))

        stats = backend.get_dashboard_stats()

        self.assertEqual(stats["project_count"], 1)
        self.assertEqual(stats["total_tasks"], len(DEFAULT_TASKS))
        self.assertEqual(stats["incomplete_count"], len(DEFAULT_TASKS))

    def test_deleting_task_removes_task_notes(self) -> None:
        backend = ProjectTrackerBackend(self.tmp / "data.json")
        project_id = backend.create_project(ProjectRecord(job_name="Job", job_number="J"))
        task = backend.list_tasks(project_id)[0]
        task_id = task.id
        assert task_id is not None

        backend.add_task_note(task_id, "note")
        backend.delete_task(task_id)

        self.assertEqual(backend.list_task_notes(task_id), [])

    def test_replacing_tasks_removes_old_task_notes(self) -> None:
        backend = ProjectTrackerBackend(self.tmp / "data.json")
        project_id = backend.create_project(ProjectRecord(job_name="Job", job_number="J"))
        old_task = backend.list_tasks(project_id)[0]
        old_task_id = old_task.id
        assert old_task_id is not None

        backend.add_task_note(old_task_id, "note")
        backend.replace_project_tasks(project_id, "phoenix")

        self.assertEqual(backend.list_task_notes(old_task_id), [])


class UpdaterRegressionTests(TempWorkspaceTest):
    def test_update_zip_requires_internal_runtime_folder(self) -> None:
        zip_path = self.tmp / "ProjectTrackingTool.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("ProjectTrackingTool.exe", "stub")

        with self.assertRaises(UpdatePackageError):
            _validate_update_zip(zip_path)

    def test_update_zip_accepts_full_flat_payload(self) -> None:
        zip_path = self.tmp / "ProjectTrackingTool.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.writestr("ProjectTrackingTool.exe", "stub")
            zf.writestr("_internal/runtime.dll", "stub")

        _validate_update_zip(zip_path)

    def test_update_script_copies_payload_contents_to_install_folder(self) -> None:
        script = _build_update_powershell_script(
            self.tmp / "ProjectTrackingTool.zip",
            self.tmp / "install",
            self.tmp / "install" / "ProjectTrackingTool.exe",
        )

        self.assertIn("Get-ChildItem -LiteralPath $payload -Force", script)
        self.assertIn("Copy-Item -Destination $installDir -Recurse -Force", script)
        self.assertIn("_internal", script)


if __name__ == "__main__":
    unittest.main()
