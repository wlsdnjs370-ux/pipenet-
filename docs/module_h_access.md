# Module H Access input circuit

H only; F keeps its landing, explicit upload confirmation and optional merge
inputs. H's native adapter delegates switching, upload, extraction and merge to
the existing engine, preserving slot identity and server-side validation.

## Interaction

Access unfolds the window horizontally (650 ms), revealing input cards in order
(110 ms stagger). Selecting a card shrinks/docks the circuit at bottom right
(760 ms), then opens that drawing's current task. Opening Access/확대 reverses the
dock; Escape folds it. Reduced-motion skips animation. Buttons lock during slot
changes and existing jobs. No upload occurs until a file is selected.

The large circuit also contains additional system slots, their existing
confirmation-based removal, and replacement upload for the current drawing.
Adding an extra system slot requires an uploaded drawing/session first; the
disabled button explains this prerequisite. A maximum of ten systems is retained.

## Completion contract

`GET /api/module-h/access-state` reads the active slot directly, and parked slots
from the session store. It never switches slots or creates/rebuilds outputs.

* Plan: nonempty node/pipe tables with no `_design_stale` result.
* System: nonempty extracted nodes/pipes for **every** system slot, including extras.
* Machine room: nonempty extracted nodes/pipes.

These mean **input material ready**, not hydraulic correctness or regulatory
approval. Native merge validation remains mandatory. Unapplied plan form values
also disable integration client-side. Failed readiness requests fail closed and
can be retried by reopening Access. No cached completion is persisted separately.
Job completion and native slot/stage changes refresh readiness; mere card visits
never mark a process complete. Replacing an H DXF clears prior extracted paths
via its explicit `h_access` upload flag, without changing F's upload behavior.

## Verification

`python -X utf8 -m pytest tests/test_module_h_access.py tests/test_module_h_access_browser.py tests/test_module_h_upload_browser.py tests/test_module_h_landing_browser.py`

Tests use isolated sessions/HTTP boundaries, never live drawings. The authenticated
read-only `scripts/check_module_h_workspace_live.py` checks deployed H and F after
an approved restart; it never uploads or opens saved work.
