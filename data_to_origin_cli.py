"""Console executable entry point; stdout remains available for JSON automation."""
import multiprocessing

from origin_bridge.cli import main

if __name__ == "__main__":
    multiprocessing.freeze_support()
    raise SystemExit(main())
