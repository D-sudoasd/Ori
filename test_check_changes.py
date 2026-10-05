from __future__ import annotations

import contextlib
import io
import json
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from scripts.check_changes import Change, changed_files, main, route_changes


class ChangeRoutingTests(unittest.TestCase):
    def test_documents_and_images_do_not_schedule_code_checks(self) -> None:
        plan = route_changes(
            [Change("README.md"), Change("docs/guide.rst"), Change("docs/flow.svg"),
             Change("docs/agent-efficiency-results.json"), Change("assets/readme/cover-source.json")]
        )
        self.assertEqual(plan["mode"], "none")
        self.assertEqual(plan["unittest_args"], [])
        self.assertEqual(plan["reasons"], ["documentation-only", "image-only"])
        self.assertFalse(plan["agent"])
        self.assertFalse(plan["build"])

    def test_test_only_change_selects_just_that_test_module(self) -> None:
        plan = route_changes([Change("test_generic_cli.py")])
        self.assertEqual(plan["unittest_args"], ["test_generic_cli"])
        self.assertFalse(plan["agent"])
        self.assertFalse(plan["build"])

    def test_leaf_modules_select_related_tests(self) -> None:
        plan = route_changes(
            [
                Change("origin_bridge/cli.py"),
                Change("origin_bridge/mcp_server.py"),
                Change("origin_bridge/gui.py"),
                Change("origin_bridge/batch.py"),
                Change("origin_bridge/exporter.py"),
            ]
        )
        self.assertEqual(
            plan["unittest_args"],
            [
                "test_batch",
                "test_batch_gui",
                "test_batch_origin",
                "test_batch_review",
                "test_generic_cli",
                "test_generic_exporter",
                "test_generic_gui",
            ],
        )
        self.assertEqual(plan["agent_unittest_args"], ["test_mcp_server"])
        self.assertTrue(plan["agent"])
        self.assertFalse(plan["build"])

    def test_exporter_covers_cli_gui_batch_and_mcp_callers(self) -> None:
        plan = route_changes([Change("origin_bridge/exporter.py")])
        self.assertEqual(
            plan["unittest_args"],
            ["test_batch", "test_generic_cli", "test_generic_exporter", "test_generic_gui"],
        )
        self.assertEqual(plan["agent_unittest_args"], ["test_mcp_server"])
        self.assertTrue(plan["agent"])
        self.assertFalse(plan["build"])

    def test_mcp_only_change_uses_agent_job_without_base_matrix(self) -> None:
        plan = route_changes([Change("origin_bridge/mcp_server.py")])
        self.assertEqual(plan["unittest_args"], [])
        self.assertEqual(plan["agent_unittest_args"], ["test_mcp_server"])
        self.assertIn("agent-tools-only", plan["reasons"])
        self.assertTrue(plan["agent"])
        self.assertFalse(plan["build"])

    def test_cli_batch_configuration_and_source_summaries_cover_callers(self) -> None:
        cases = {
            "origin_bridge/cli.py": ["test_batch_review", "test_generic_cli"],
            "origin_bridge/batch_config.py": ["test_batch_gui"],
            "origin_bridge/source_summary.py": ["test_generic_gui"],
        }
        for path, expected in cases.items():
            with self.subTest(path=path):
                self.assertEqual(route_changes([Change(path)])["unittest_args"], expected)

    def test_packaged_worker_verifier_runs_build_without_full_suite(self) -> None:
        plan = route_changes([Change("scripts/verify_packaged_worker.py")])
        self.assertEqual(plan["unittest_args"], [])
        self.assertEqual(plan["mode"], "build-only")
        self.assertFalse(plan["agent"])
        self.assertTrue(plan["build"])

    def test_shared_core_including_new_source_cache_selects_full_path(self) -> None:
        plan = route_changes(
            [
                Change("origin_bridge/readers.py"),
                Change("origin_bridge/planning.py"),
                Change("origin_bridge/models.py"),
                Change("origin_bridge/dates.py"),
                Change("origin_bridge/source_cache.py", "??"),
            ]
        )
        self.assertEqual(plan["unittest_args"], ["discover", "-v"])
        self.assertTrue(plan["agent"])
        self.assertTrue(plan["build"])

    def test_new_module_and_unknown_configuration_are_conservative(self) -> None:
        for change in (
            Change("origin_bridge/new_format.py", "A"),
            Change("tools/runtime_hook.py", "??"),
            Change("custom-build.toml"),
        ):
            with self.subTest(path=change.path):
                plan = route_changes([change])
                self.assertEqual(plan["unittest_args"], ["discover", "-v"])
                self.assertTrue(plan["build"])
                self.assertTrue(plan["agent"])

    def test_shared_runtime_requirements_run_all_client_paths(self) -> None:
        plan = route_changes([Change("requirements.txt")])
        self.assertEqual(plan["unittest_args"], ["discover", "-v"])
        self.assertTrue(plan["agent"])
        self.assertTrue(plan["build"])

    def test_workflow_configuration_is_a_full_agent_and_build_boundary(self) -> None:
        plan = route_changes([Change(".github/workflows/tests.yml")])
        self.assertEqual(plan["unittest_args"], ["discover", "-v"])
        self.assertTrue(plan["agent"])
        self.assertTrue(plan["build"])

    def test_wrappers_workers_and_session_require_a_build(self) -> None:
        for path in (
            "data_to_origin_cli.py",
            "origin_bridge/worker.py",
            "origin_bridge/session.py",
            "SpectraToOrigin.spec",
        ):
            with self.subTest(path=path):
                plan = route_changes([Change(path)])
                self.assertEqual(plan["unittest_args"], ["discover", "-v"])
                self.assertTrue(plan["build"])

    def test_test_deletion_runs_remaining_suite_without_packaging(self) -> None:
        plan = route_changes([Change("test_generic_cli.py", "D")])
        self.assertEqual(plan["unittest_args"], ["discover", "-v"])
        self.assertFalse(plan["build"])

    def test_explicit_full_and_empty_base_fallback_are_deep_paths(self) -> None:
        forced = route_changes([], full=True)
        fallback = route_changes([], fallback="missing-or-new-push-base")
        for plan in (forced, fallback):
            with self.subTest(mode=plan["mode"]):
                self.assertEqual(plan["unittest_args"], ["discover", "-v"])
                self.assertTrue(plan["agent"])
                self.assertTrue(plan["build"])

        for base in ("", "0" * 40):
            output = io.StringIO()
            with contextlib.redirect_stdout(output):
                self.assertEqual(main(["--base", base]), 0)
            self.assertEqual(json.loads(output.getvalue())["mode"], "full")

    def test_run_uses_this_python_and_github_outputs_are_machine_readable(self) -> None:
        invoked: list[list[str]] = []

        def fake_run(command: list[str], **_: object) -> subprocess.CompletedProcess[bytes]:
            invoked.append(command)
            return subprocess.CompletedProcess(command, 0)

        with tempfile.TemporaryDirectory() as directory:
            output_path = Path(directory) / "outputs"
            output = io.StringIO()
            with (
                mock.patch("scripts.check_changes.changed_files", return_value=[Change("test_generic_cli.py")]),
                mock.patch("scripts.check_changes.sys.executable", "same-python"),
                mock.patch("scripts.check_changes.subprocess.run", side_effect=fake_run),
                contextlib.redirect_stdout(output),
            ):
                self.assertEqual(main(["--run", f"--github-output={output_path}"]), 0)
            self.assertEqual(invoked, [["same-python", "-m", "unittest", "test_generic_cli"]])
            self.assertEqual(
                output_path.read_text(encoding="utf-8"),
                "unittest_args=test_generic_cli\nagent=false\nbuild=false\n",
            )
            self.assertEqual(json.loads(output.getvalue())["unittest_args"], ["test_generic_cli"])

    def test_run_executes_selected_agent_check_with_the_same_python(self) -> None:
        invoked: list[list[str]] = []

        def fake_run(command: list[str], **_: object) -> subprocess.CompletedProcess[bytes]:
            invoked.append(command)
            return subprocess.CompletedProcess(command, 0)

        with (
            mock.patch("scripts.check_changes.changed_files", return_value=[Change("origin_bridge/mcp_server.py")]),
            mock.patch("scripts.check_changes.sys.executable", "same-python"),
            mock.patch("scripts.check_changes.subprocess.run", side_effect=fake_run),
            contextlib.redirect_stdout(io.StringIO()),
        ):
            self.assertEqual(main(["--run"]), 0)
        self.assertEqual(invoked, [["same-python", "-m", "unittest", "test_mcp_server"]])

    def test_changed_files_include_staged_unstaged_and_nonignored_untracked(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._git(root, "init", "-q")
            self._git(root, "config", "user.email", "route@example.test")
            self._git(root, "config", "user.name", "Route Test")
            (root / ".gitignore").write_text("*.ignored\n", encoding="utf-8")
            (root / "base.txt").write_text("base\n", encoding="utf-8")
            self._git(root, "add", ".gitignore", "base.txt")
            self._git(root, "commit", "-qm", "base")
            base = subprocess.run(
                ["git", "rev-parse", "HEAD"], cwd=root, check=True,
                capture_output=True, text=True,
            ).stdout.strip()

            (root / "base.txt").write_text("changed\n", encoding="utf-8")
            (root / "staged.py").write_text("# staged\n", encoding="utf-8")
            self._git(root, "add", "staged.py")
            (root / "new.py").write_text("# untracked\n", encoding="utf-8")
            (root / "ignored.ignored").write_text("ignored\n", encoding="utf-8")

            for changes in (changed_files(cwd=root), changed_files(base, cwd=root)):
                with self.subTest(changes=changes):
                    self.assertEqual(
                        {change.path: change.status for change in changes},
                        {"base.txt": "M", "staged.py": "A", "new.py": "??"},
                    )

    @staticmethod
    def _git(root: Path, *args: str) -> None:
        subprocess.run(["git", *args], cwd=root, check=True, capture_output=True)


if __name__ == "__main__":
    unittest.main()
