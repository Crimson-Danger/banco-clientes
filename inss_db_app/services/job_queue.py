from __future__ import annotations

import logging
import threading
from typing import Any, Callable

from ..config import settings


def _enqueue_with_thread(target: Callable[..., Any], args: tuple[Any, ...], daemon: bool = True) -> str:
    threading.Thread(target=target, args=args, daemon=daemon).start()
    return "local-thread"


def _enqueue_with_rq(queue_name: str, callable_path: str, kwargs: dict[str, Any]) -> str:
    from redis import Redis
    from rq import Queue

    redis_conn = Redis.from_url(settings.queue_redis_url)
    queue = Queue(name=queue_name or settings.queue_default_name, connection=redis_conn)
    job_timeout = kwargs.pop("_rq_job_timeout", None)
    enqueue_kwargs: dict[str, Any] = {"kwargs": kwargs}
    if job_timeout is not None:
        enqueue_kwargs["job_timeout"] = int(job_timeout)
    job = queue.enqueue(callable_path, **enqueue_kwargs)
    return str(job.id)


def enqueue_background_job(
    *,
    queue_name: str,
    callable_path: str,
    fallback_target: Callable[..., Any],
    fallback_args: tuple[Any, ...],
    fallback_daemon: bool = True,
    kwargs: dict[str, Any] | None = None,
    rq_job_timeout: int | None = None,
) -> dict[str, str]:
    backend = settings.queue_backend
    payload = dict(kwargs or {})
    if rq_job_timeout is not None:
        payload["_rq_job_timeout"] = int(rq_job_timeout)
    if backend == "rq":
        try:
            queue_job_id = _enqueue_with_rq(queue_name=queue_name, callable_path=callable_path, kwargs=payload)
            return {"enqueue_mode": "rq", "queue_job_id": queue_job_id, "queue_name": queue_name}
        except Exception as exc:
            logging.getLogger("inss_app").warning(
                "Falha ao enfileirar em RQ, usando thread local. erro=%s", str(exc)
            )
    local_id = _enqueue_with_thread(target=fallback_target, args=fallback_args, daemon=fallback_daemon)
    return {"enqueue_mode": "thread", "queue_job_id": local_id, "queue_name": queue_name}


def cancel_rq_job(queue_job_id: str) -> bool:
    if not queue_job_id:
        return False
    from redis import Redis
    from rq.command import send_stop_job_command
    from rq.job import Job

    redis_conn = Redis.from_url(settings.queue_redis_url)
    job = Job.fetch(queue_job_id, connection=redis_conn)
    if job.is_finished:
        return False
    if job.is_started:
        send_stop_job_command(redis_conn, queue_job_id)
    job.cancel()
    return True
