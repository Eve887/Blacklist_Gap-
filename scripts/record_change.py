"""
record_change.py - backup + diff helper for PROJECT_RULES.md rule B.

Standard library only (shutil, difflib) - nothing to install.

As a library (the normal way this gets used, from another script or an
interactive Python call):

    from record_change import before_edit, after_edit

    original = before_edit("scripts/common.py")   # call BEFORE editing
    ... make the edit ...
    before_path, diff_path = after_edit("scripts/common.py", original)

`before_edit` returns None for a file that does not exist yet (a NEW file).
`after_edit` then saves a ".created" copy instead of a ".before"/".diff"
pair. For a file being DELETED, call `after_edit(path, original, deleted=True)`
AFTER taking the snapshot but BEFORE actually deleting the file on disk (it
needs to still read the "before" text; the deleted copy is identical to
that, just named ".deleted" instead of ".before").

As a one-off command line tool, for a quick manual backup without writing
a diff (e.g. before an edit you'll do by hand, not through a script):

    py scripts\\record_change.py before scripts\\common.py
    ... now edit the file ...
    py scripts\\record_change.py after scripts\\common.py

Everything is written to code_history\\<YYYY-MM-DD_HHMM>_<filename>.before
/.diff /.created /.deleted, as PROJECT_RULES.md rule B requires.
"""

import sys
import shutil
import difflib
from datetime import datetime
from pathlib import Path

PROJECT = Path(__file__).resolve().parent.parent
HISTORY = PROJECT / "code_history"
HISTORY.mkdir(parents=True, exist_ok=True)

_PENDING_DIR = HISTORY / ".pending"
_PENDING_DIR.mkdir(parents=True, exist_ok=True)


def _stamp():
    return datetime.now().strftime("%Y-%m-%d_%H%M")


def before_edit(path):
    """Call BEFORE editing `path`. Returns the file's current text, or
    None if the file does not exist yet (a new file)."""
    path = Path(path)
    if not path.exists():
        return None
    return path.read_text(encoding="utf-8", errors="replace")


def after_edit(path, original_text, deleted=False):
    """Call AFTER editing (or right before actually deleting) `path`.
    Writes the code_history file(s) and returns (first_path, diff_path) -
    diff_path is None for a new or deleted file (nothing to diff against)."""
    path = Path(path)
    base = HISTORY / f"{_stamp()}_{path.name}"

    if original_text is None:
        created = Path(str(base) + ".created")
        shutil.copy(path, created)
        return created, None

    if deleted:
        deleted_path = Path(str(base) + ".deleted")
        deleted_path.write_text(original_text, encoding="utf-8")
        return deleted_path, None

    before_path = Path(str(base) + ".before")
    before_path.write_text(original_text, encoding="utf-8")

    new_text = path.read_text(encoding="utf-8", errors="replace")
    diff_path = Path(str(base) + ".diff")
    diff_lines = difflib.unified_diff(
        original_text.splitlines(keepends=True),
        new_text.splitlines(keepends=True),
        fromfile=f"{path.name} (before)",
        tofile=f"{path.name} (after)",
    )
    diff_path.write_text("".join(diff_lines), encoding="utf-8")
    return before_path, diff_path


def _pending_file(path):
    return _PENDING_DIR / (Path(path).name + ".pending")


def main():
    if len(sys.argv) != 3 or sys.argv[1] not in ("before", "after", "deleted"):
        print(__doc__)
        sys.exit(1)
    action, path = sys.argv[1], Path(sys.argv[2])
    pending = _pending_file(path)

    if action == "before":
        text = before_edit(path)
        pending.write_text("\x00NEW\x00" if text is None else text, encoding="utf-8")
        print(f"Snapshot taken for {path}")
    else:
        if not pending.exists():
            print(f"No 'before' snapshot found for {path} - run 'before' first.")
            sys.exit(1)
        saved = pending.read_text(encoding="utf-8")
        original_text = None if saved == "\x00NEW\x00" else saved
        before_path, diff_path = after_edit(path, original_text, deleted=(action == "deleted"))
        print(f"Saved: {before_path}")
        if diff_path:
            print(f"Saved: {diff_path}")
        pending.unlink()


if __name__ == "__main__":
    main()
