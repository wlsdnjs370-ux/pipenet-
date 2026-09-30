# Module H process navigation

The footer's `돌아가기` button sits immediately before `처리 내역`.
Workflow navigation is button-only. Ctrl+Z (also Meta+Z) instead uses the shared
engine's existing editing history: undo the last edit in the current stage,
without changing stages. Repeated presses undo earlier edits one at a time;
an empty history stays in the current stage. Ctrl+Shift+Z retains the existing
network-editor redo behavior in attribute definition and integration.

The redundant canvas coordinate/status strip is no longer displayed in H.
The original `화면 맞춤` control is moved into the footer before `돌아가기`,
retaining its camera handler and availability during background processing.
Status/error anchors remain for the processing/history UI; Module F is unchanged.

- Return to the previous reachable native workflow stage, retaining existing
  stage-entry validation and data. This does not delete output or undo graph edits.
- From integration, return to the stage from which integration was entered.
  If unavailable, use the last reachable stage in the current drawing's workflow.
- Drawing/session/method changes discard the integration return destination.
- Upload is the first stage: no further return is available.
- Block navigation during processing, pending navigation, and modal dialogs.
- Block editing undo/redo during processing and modal dialogs; held-key repeats
  cannot consume several edits. Native text undo remains browser-owned.
- Preserve native input/text undo and the existing explicit graph-undo controls.
- Module F's interface, shortcuts and calculation rules are unchanged.

`tests/test_module_h_back_browser.py` covers stage sequences, integration return,
system/machine-room and automatic workflows, data preservation, button placement,
native text undo, repeated keys, processing and modal guards in an isolated server.
`tests/test_module_h_undo_browser.py` covers one-edit-at-a-time keyboard undo/redo
of real network properties in plan and integration, native per-stage undo routing,
and empty-history behavior without workflow navigation.
