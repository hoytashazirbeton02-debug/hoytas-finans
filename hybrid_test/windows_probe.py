import os
import sys
import getpass
from pathlib import Path
from offline_queue import OfflineQueue
from transport import send_event
from sync_core import flush

def main():
    directory = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "HOYTAS-Hybrid-Test"
    directory.mkdir(parents=True, exist_ok=True)
    queue = OfflineQueue(directory / "probe.sqlite")
    if "--offline" in sys.argv:
        print("Queued:", queue.enqueue("HOYTAS_SYNTHETIC_TEST"))
        return
    token = getpass.getpass("Test API token: ")
    if not queue.pending():
        queue.enqueue("HOYTAS_SYNTHETIC_TEST")
    flush(queue, lambda eid, payload: send_event("https://hoytas-hibrit-api-test.onrender.com", token, eid, payload))
    print("Remaining:", len(queue.pending()))

if __name__ == "__main__":
    main()
