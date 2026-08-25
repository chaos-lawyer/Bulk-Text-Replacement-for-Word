from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
import sys
sys.path.insert(0, str(ROOT / "src"))

from application.task_coordinator import TaskCoordinator
from application.workers import TaskWorker


class TaskCoordinatorTests(unittest.TestCase):
    def setUp(self):
        self.coordinator = TaskCoordinator()

    def test_single_write_task_exclusivity(self):
        worker1 = TaskWorker(lambda: None)
        worker2 = TaskWorker(lambda: None)

        self.assertTrue(self.coordinator.can_start_task(is_write=True))
        task_id = self.coordinator.register_task(worker1, is_write=True)
        self.assertTrue(self.coordinator.is_write_busy)
        self.assertFalse(self.coordinator.can_start_task(is_write=True))
        self.assertFalse(self.coordinator.can_start_task(is_write=False))

        # Attempting to register another write task raises RuntimeError
        with self.assertRaises(RuntimeError):
            self.coordinator.register_task(worker2, is_write=True)

        # Release task with matching ID
        self.assertTrue(self.coordinator.release_task(task_id))
        self.assertFalse(self.coordinator.is_write_busy)
        self.assertTrue(self.coordinator.can_start_task(is_write=True))

    def test_stale_task_id_does_not_release(self):
        worker = TaskWorker(lambda: None)
        task_id = self.coordinator.register_task(worker, is_write=True)

        # Wrong task ID cannot release
        self.assertFalse(self.coordinator.release_task(task_id + 99))
        self.assertTrue(self.coordinator.is_write_busy)

    def test_cancel_active_task(self):
        worker = TaskWorker(lambda: None)
        task_id = self.coordinator.register_task(worker, is_write=False, description="扫描任务")
        self.assertEqual(self.coordinator.active_task_description, "扫描任务")
        self.assertTrue(self.coordinator.cancel_active_task())
        self.assertTrue(worker.cancel_token.is_cancelled)
        self.coordinator.release_task(task_id)
        self.assertFalse(self.coordinator.is_busy)
