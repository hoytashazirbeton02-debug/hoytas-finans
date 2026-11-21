"""Start the HOYTAŞ shared Windows profile with the main application's central login."""
import os
from pathlib import Path
import runpy

from hybrid_windows_profile import ENDPOINT


if __name__ == "__main__":
    profile = Path(os.environ["LOCALAPPDATA"]) / "HOYTASFinans" / "canli-kullanim"
    os.environ.update(HOYTAS_PROFILE_DIR=str(profile), HOYTAS_HYBRID_CAPTURE="1",
                      HOYTAS_SYNC_URL=ENDPOINT, HOYTAS_MEMBERSHIP_URL=ENDPOINT)
    os.environ.pop("HOYTAS_SYNC_TOKEN", None)
    runpy.run_path(str(Path(__file__).with_name("hoytas.py")), run_name="__main__")
