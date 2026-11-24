"""Start the optional desktop sync worker after an authenticated local login."""
from datetime import datetime
import os
from pathlib import Path
from queue import Empty, Queue
from threading import Thread
from time import monotonic
from urllib.error import HTTPError
from urllib.parse import urlsplit

from hybrid_exchange import exchange_once
from hybrid_http_relay import HTTPRelay
from hybrid_identity import replica_slot
from hybrid_inbound import ReplicaConflict

TEST_HOSTS = {"hoytas-aktarim-pilot.onrender.com", "hoytas-hibrit-api-test.onrender.com"}


def production_relay(connect, endpoint, token):
    if os.environ.get("HOYTAS_HYBRID_CAPTURE") != "1":
        raise RuntimeError("Hibrit oturum ayarı kapalı.")
    if urlsplit(endpoint).hostname in TEST_HOSTS:
        raise ValueError("İşletme kayıtları deneme sunucusuna gönderilemez.")
    relay = HTTPRelay(endpoint, token)
    with connect() as db:
        if replica_slot(db) is None:
            raise RuntimeError("Bu veritabanı için hibrit oturum kimliği oluşturulmamış.")
    return relay


def start_background_sync(app, connect, data_directory):
    endpoint = os.environ.get("HOYTAS_SYNC_URL", "").strip()
    if not endpoint:
        return
    session = getattr(app, "membership_session", None)
    token = session or os.environ.get("HOYTAS_SYNC_TOKEN", "")
    relay = production_relay(connect, endpoint, token)
    results = Queue()
    client = getattr(app, "membership_client", None)
    with connect() as db:
        primary = replica_slot(db) == 1
    publish_pending = session and client is not None and app.current_user["role"] == "Yönetici" and primary
    restore_pending = session and client is not None and not primary
    publish_retry_at = 0

    def run():
        nonlocal publish_pending, publish_retry_at, restore_pending
        restored = False
        if restore_pending:
            restore_pending = False
            from hybrid_enrollment import restore_enrolled_history

            try:
                restored = restore_enrolled_history(Path(data_directory) / "hoytas.db", client, session, connect)
            except Exception as error:
                app._log_update_error(error)
        try:
            result = exchange_once(connect, relay, limit=5)
            if publish_pending and result == (0, 0) and monotonic() >= publish_retry_at:
                from hybrid_enrollment import read_baseline

                try:
                    baseline = read_baseline(Path(data_directory) / "hoytas.db")
                    client.publish_baseline(token, baseline)
                    publish_pending = False
                except (OSError, ValueError, RuntimeError) as error:
                    app._log_update_error(error)
                    publish_retry_at = monotonic() + 60
            results.put((result, None, restored))
        except Exception as error:
            results.put((None, error, restored))

    def poll():
        try:
            result, error, restored = results.get_nowait()
        except Empty:
            app.after(250, poll)
            return

        if restored or (result is not None and result[0]):
            app.refresh_all()
        if error is not None:
            try:
                with (Path(data_directory) / "senkronizasyon.log").open("a", encoding="utf-8") as handle:
                    handle.write(f"{datetime.now().isoformat(timespec='seconds')} {type(error).__name__}\n")
            except OSError:
                pass
            if isinstance(error, HTTPError) and error.code == 401 and getattr(app, "membership_session", None):
                app.title("HOYTAŞ — Oturum sona erdi, yeniden giriş yapın")
                return
            if isinstance(error, ReplicaConflict):
                app.title("HOYTAŞ — Senkronizasyon çakışması")
                return
            app.title("HOYTAŞ — Sunucu bağlantısı bekleniyor")
            app.after(30000, launch)
            return

        received, sent = result
        app.title(original_title)
        app.after(250 if received == 5 or sent == 5 else 5000, launch)

    def launch():
        Thread(target=run, name="HOYTAS-Sync", daemon=True).start()
        app.after(250, poll)

    original_title = app.title()
    app.after(100, launch)
