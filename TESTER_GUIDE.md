# CBMI Loop — Beta Tester Guide

Thanks for testing CBMI Loop. It runs fully on your machine — no internet
required, your survey data never leaves your computer.

> **Beta builds expire after 14 days.** When yours expires, email
> **bhanu@crystalball.ai** for a fresh build.

---

## Install

### macOS (.dmg)

1. Open the `.dmg` and drag **CBMI Loop** into **Applications**.
2. The first launch is blocked because the app isn't signed with an Apple
   certificate (expected for a beta). You'll see *"Apple could not verify…"* or
   *"…is damaged"*.
3. Allow it **once**:
   - Open **System Settings → Privacy & Security**.
   - Scroll to the **Security** section — you'll see *"CBMI Loop was blocked"*.
   - Click **Open Anyway**, then confirm with **Open**.
   - *(If the app was quarantined, you can instead run once in Terminal:*
     `xattr -dr com.apple.quarantine "/Applications/CBMI Loop.app"` *)*
4. After that, launch it normally from Applications.

### Windows (Setup.exe)

1. Run `CBMI-Loop-…-Setup.exe`.
2. Windows SmartScreen will warn *"Windows protected your PC"* (expected — the
   installer isn't code-signed for beta). Click **More info → Run anyway**.
3. Follow the installer; launch **CBMI Loop** from the Start menu or desktop.
4. If Windows Defender or your AV quarantines it, restore/allow it — frozen
   Python apps sometimes trigger false positives.

---

## Using it

1. On the **Load** page, pick a subsystem tab (Drone / Base / Control Point /
   Check Point). Required files are marked with a red **\***; **optional** ones
   are labelled.
2. Add your files, then click **Validate & Run Scoring**. A progress panel
   shows each subsystem moving through Validate → Process & Score →
   Recommendations → Excel → Provenance.
3. When it finishes, click **View on hardware page** to see the score, building
   blocks, indicators, and recommendations — or **Download Excel**.
4. Past runs appear under **Previous runs** on the Load page; click **Open
   results** to reopen any of them.

Large drone datasets (RINEX parsing) can take several minutes — that's normal.

---

## Sending feedback

If something breaks, click **Export diagnostics** (bottom of the Load page)
and email the downloaded `cbmi-diagnostics.zip` to **bhanu@crystalball.ai**.
It contains app logs and version info only — **no survey data and no scoring
internals**.
