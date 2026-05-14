from __future__ import annotations

import os

from inss_db_app.config import settings


def main() -> None:
    if settings.queue_backend != "rq":
        print(
            "RQ worker nao iniciado: APP_QUEUE_BACKEND!="
            f"rq (valor atual: {settings.queue_backend})."
        )
        return

    from redis import Redis
    from rq import Connection, SimpleWorker, Worker
    from rq.timeouts import TimerDeathPenalty

    queue_names = [settings.queue_exports_name, settings.queue_imports_name, settings.queue_default_name]
    queue_names = [name for idx, name in enumerate(queue_names) if name and name not in queue_names[:idx]]

    print(
        "Startup RQ worker:",
        f"redis={settings.queue_redis_url}",
        f"queues={','.join(queue_names)}",
    )

    redis_conn = Redis.from_url(settings.queue_redis_url)
    with Connection(redis_conn):
        worker_cls = SimpleWorker if os.name == "nt" else Worker
        worker = worker_cls(queue_names)
        if os.name == "nt":
            worker.death_penalty_class = TimerDeathPenalty
        worker.work(with_scheduler=False)


if __name__ == "__main__":
    main()
