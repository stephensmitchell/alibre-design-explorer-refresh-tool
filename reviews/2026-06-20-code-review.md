# alibre-design-explorer-refresh-tool — Code Review (Correctness)

**Date:** 2026-06-20
**Scope:** Second-opinion review, code only (correctness bugs).

**Summary: No High or Medium correctness bugs found. 1 Low-severity resource leak.**

The only source file with meaningful logic is `source/Init.vb`. The COM startup, null/type guards (`Is Nothing OrElse Not TypeOf session Is IADPartSession`), exception handling, and reverse-order COM release in the `Finally` block are all correct.

## Low

- **source/Init.vb:31-34** — The `IADPartFeature` objects yielded by the `For Each partSession.Features` loop (and the `Features` collection RCW from lines 28 and 31) are never released via `Marshal.ReleaseComObject`, unlike the other COM objects that are explicitly released in the `Finally` block. This is an inconsistent COM RCW leak; it is low severity only because the process is short-lived and exits immediately, so process teardown reclaims the references.

No logic errors, off-by-one, wrong-operator, null-dereference, control-flow, or error-swallowing bugs were found. The `GetObject(, "AlibreX.AutomationHook")` call correctly fetches a running instance, and its `COMException` handler correctly reports the "not running" case.

---

## Fixes applied — 2026-06-20

- **[Low] `source/Init.vb`** — the `Features` collection is now captured once and released via `Marshal.ReleaseComObject`, and each `IADPartFeature` RCW is released at the end of the loop; the collection is released first in the `Finally` block (reverse order, matching the existing idiom).

*Caveat: changes applied to source; not verified by build. The `IADPartFeatures` collection type could not be confirmed against local stubs (none present), but matches the standard AlibreX interface.*
