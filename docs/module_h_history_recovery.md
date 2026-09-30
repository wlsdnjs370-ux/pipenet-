# Module H edit history and changed drawing bases

Ctrl+Z undoes the latest edit in the current stage, not a workflow transition.
The separate `돌아가기` button remains process navigation. Native text undo is
unchanged; busy operations, modal dialogs and held-key repeats cannot consume
graph history. Existing network-editor redo remains Ctrl+Shift+Z.

The network-definition drawer no longer displays the duplicate
`영역 배관망 다시 추출` row or its modification counter. The ordinary
`영역 배관망 추출` action still handles both initial and updated selections.

## Changed-base recovery

A saved edit addresses an exact graph fingerprint, not merely a pipe label.
Re-extracting a different selection can reuse a label such as P71 for another
pipe. Replaying the old property edit would silently alter the wrong pipe.
Identical bases retain their existing replay/undo behavior.

When a changed basis prevents input preview, H now refreshes editor state even
if no property table is open. An independent recovery panel exposes
`이전 편집 보관 후 계속`. It appears only for a
conflict in the current drawing and requires confirmation explaining that the
old edits are preserved separately, not applied to the new graph.

`accept_basis` is accepted only for a current design conflict and exact revision.
It also checks selection freshness, processing state and concurrent disk writes.
The original history bytes are archived before the active history changes; a
backup/save failure leaves the graph and active history intact. No commands are
mapped across different bases. Prior branches remain available to the existing
exact-base restoration mechanism. The integrated graph is not silently rebuilt.

No actual user history is reset or migrated as part of deploying this change.
Backend changes require a server restart; active in-memory drawings must be
saved before restarting. Tests use isolated sessions, graphs and history files.

## Visible review gate (2026-09-29)

The processing-history drawer hid the attributes toolbar with CSS visibility,
including its only recovery button. Recovery now has an independent, movable
panel, visible even while that drawer is open. The primary action leads to this
review rather than repeating the build; the automatic attributes rebuild waits
for the decision. All checks are scoped to the current session, plan and stage.
If the selection has changed again, a read-only freshness check disables basis
acceptance and offers `현재 선택으로 기준망 갱신` instead of a dead-end retry.

Only an exact current-editor conflict message is labelled as review required.
Unrelated errors and server job failures remain failures. Preview and export
guards still reject conflicts; no history matching tolerance, graph mutation,
automatic acceptance or automatic migration is introduced. The existing explicit
confirmation, byte-for-byte archival and revision/disk checks remain authoritative.
