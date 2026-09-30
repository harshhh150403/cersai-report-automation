"""Command line for the three steps.

    python -m cersai_tool run1            build the working Excel + required-details file
    python -m cersai_tool run2            apply the required details -> final Excel
    python -m cersai_tool run3            check the final Excel and write the .DAT file

Options:
    --batch 2026-09-30   work on this batch (default: the latest suitable one; you are asked to confirm)
    --yes                do not ask for confirmation
    --dummy              use "Dummy input data" and write to "Dummy output data" (for testing)
    --root PATH          the project folder (default: the folder that contains cersai_tool)
"""
import argparse
import sys
import traceback
from datetime import datetime

from .batch import BatchError
from .settings import make_paths
from .sources.lms import InputError
from .sources.required_details import RequiredDetailsError
from .steps import step1_build, step2_apply, step3_dat


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    parser = argparse.ArgumentParser(prog="cersai_tool", description="CERSAI SI registration batch tool")
    parser.add_argument("step", choices=["run1", "run2", "run3"],
                        help="run1: build the Excel files; run2: apply the required details; run3: write the .DAT")
    parser.add_argument("--batch", help="batch folder name, e.g. 2026-09-30 (default: the latest)")
    parser.add_argument("--yes", action="store_true", help="do not ask for confirmation")
    parser.add_argument("--dummy", action="store_true", help="use the dummy input data (for testing)")
    parser.add_argument("--root", help="project folder (default: the folder that contains cersai_tool)")
    args = parser.parse_args(argv)
    paths = make_paths(args.root, args.dummy)
    if args.dummy:
        print("Using the DUMMY input data")
    try:
        if args.step == "run1":
            step1_build.run(paths)
        elif args.step == "run2":
            step2_apply.run(paths, args.batch, args.yes)
        else:
            step3_dat.run(paths, args.batch, args.yes)
    except PermissionError as error:
        name = getattr(error, "filename", "") or "a file"
        print(f"\nERROR: could not open or save {name}. If it is open in Excel, close it and run again.")
        return 2
    except (InputError, BatchError, RequiredDetailsError, FileNotFoundError) as error:
        print(f"\nERROR: {error}")
        return 1
    except KeyboardInterrupt:
        print("\nStopped.")
        return 130
    except Exception:
        log = paths.output_folder / "error log.txt"
        try:
            log.parent.mkdir(parents=True, exist_ok=True)
            with log.open("a", encoding="utf-8") as handle:
                handle.write(f"\n=== {datetime.now():%Y-%m-%d %H:%M:%S} {args.step}\n{traceback.format_exc()}")
            where = f" Details were saved in {log}."
        except OSError:
            where = ""
        print(f"\nERROR: something unexpected went wrong.{where}")
        traceback.print_exc()
        return 1
    return 0
