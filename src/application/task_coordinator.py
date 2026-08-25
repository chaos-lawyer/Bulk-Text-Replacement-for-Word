"""Centralized Task Coordinator to manage exclusive task locks and avoid re-entrant task execution."""

from __future__ import annotations

import threading
from typing import Optional

from application.workers import TaskWorker


class TaskCoordinator:
    """Manages application-wide task concurrency and worker lifecycle.

    Guarantees:
    1. Single active task policy: only ONE background task (scan, preview, or write) runs at a time.
    2. Prevents re-entrant invocations via buttons, menus, and shortcuts (Ctrl+R, etc.).
    3. Safe task release: worker finish signals only release their own matching task ID.
    4. Safe cancellation: bottom status bar and window close events route directly to active task.
    """

    _instance: Optional[TaskCoordinator] = None
    _lock = threading.Lock()

    def __init__(self):
        self._task_id_counter = 0
        self._active_task_id: Optional[int] = None
        self._active_worker: Optional[TaskWorker] = None
        self._active_desc: str = ""
        self._is_write_task: bool = False
        self._state_lock = threading.Lock()

    @classmethod
    def instance(cls) -> TaskCoordinator:
        with cls._lock:
            if cls._instance is None:
                cls._instance = TaskCoordinator()
            return cls._instance

    @property
    def is_busy(self) -> bool:
        with self._state_lock:
            return self._active_task_id is not None

    @property
    def is_write_busy(self) -> bool:
        with self._state_lock:
            return self._active_task_id is not None and self._is_write_task

    @property
    def active_task_description(self) -> str:
        with self._state_lock:
            return self._active_desc

    def can_start_task(self, is_write: bool = False) -> bool:
        """Check whether a new background task is permitted to start."""
        with self._state_lock:
            return self._active_task_id is None

    def register_task(self, worker: TaskWorker, is_write: bool = False, description: str = "") -> int:
        """Register a new task worker and obtain a unique task ID.

        Raises RuntimeError if another background task is already executing.
        """
        with self._state_lock:
            if self._active_task_id is not None:
                current = self._active_desc or "后台任务"
                raise RuntimeError(f"已有任务【{current}】正在执行中，请等待完成或取消后再试。")

            self._task_id_counter += 1
            task_id = self._task_id_counter
            self._active_task_id = task_id
            self._active_worker = worker
            self._active_desc = description
            self._is_write_task = is_write
            return task_id

    def release_task(self, task_id: int) -> bool:
        """Release the active task lock if the given task_id matches the current active task."""
        with self._state_lock:
            if self._active_task_id == task_id:
                self._active_task_id = None
                self._active_worker = None
                self._active_desc = ""
                self._is_write_task = False
                return True
            return False

    def cancel_active_task(self) -> bool:
        """Cancel the current active worker if any."""
        with self._state_lock:
            if self._active_worker is not None:
                self._active_worker.cancel()
                return True
            return False
