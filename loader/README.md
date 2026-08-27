# Keeping everyone up to date from SharePoint

An add-in cannot replace itself while Excel has it open, so "download the new
version over the top on startup" does not work. This gets round it by never
installing the real add-in at all.

```
  SharePoint (synced)          each machine
  ------------------           ------------
  WorkbookDoctor.xlam   --->   %LOCALAPPDATA%\WorkbookDoctor\WorkbookDoctor.xlam
  the master copy              the cache, which is what Excel actually runs
                                          ^
                                          |
                               WorkbookDoctorLoader.xlam
                               installed once, rarely changes
```

On Excel start the loader compares the master against the cache, copies it down
if it is newer, and opens the cache. Because the master is never the open file,
**you can replace it whenever you like** — including while people have Excel
running. They pick it up next time they start, or immediately via
**Add-ins → Workbook Doctor Updates → Check for updates now**, which reloads
the add-in without restarting Excel.

Anyone off the network keeps working from their cache.

## Setting it up

**1. Pick the SharePoint folder** and have everyone sync it (Documents library →
**Sync**). It then appears as an ordinary local path, something like:

```
C:\Users\<them>\Alpha FMC\Tools - Documents\Workbook Doctor\
```

**2. Edit one line** in [`modLoader.bas`](modLoader.bas) — everything after the
user's own folder:

```vba
Private Const SOURCE_RELATIVE As String = "Alpha FMC\Tools - Documents\Workbook Doctor\WorkbookDoctor.xlam"
```

The `%USERPROFILE%\` part is filled in per machine, which is why one constant
works for everybody. Anyone whose path differs can point their own copy at it
with **Change the shared folder**, and it is remembered.

**3. Build the master** and put it in that folder:

- VBA route: import `build\BuildAddIn.bas`, run **`BuildMasterForSharing`**.
- PowerShell: `.\Build-AddIn.ps1` and copy `dist\WorkbookDoctor.xlam` across.

**4. Build the loader** and give it to people:

- VBA route: run **`BuildLoader`** — builds and installs it here.
- PowerShell: `.\Build-AddIn.ps1 -Loader -Install`

Distribute `WorkbookDoctorLoader.xlam` with the scripts in
[`../install`](../install), renaming it in the `.cmd` if you use those.

## Releasing an update afterwards

Rebuild the master with `BuildMasterForSharing` and drop it in the SharePoint
folder. That is the whole release process. Nobody reinstalls anything.

## What people see

A **Workbook Doctor Updates** menu on the Add-ins tab, next to the add-in's own:

| | |
| --- | --- |
| Check for updates now | Fetches and reloads without restarting Excel |
| Change the shared folder | Points this machine at the master and remembers it |
| Where is it loading from? | Both paths and both dates — the first thing to look at when someone says their menu has gone |

## Things worth knowing

- **Files On-Demand.** If the library is synced but not downloaded, the master
  is a placeholder. Reading its date works; copying it triggers a download,
  which is slow the first time. Marking the folder *Always keep on this device*
  avoids the pause.
- **Timestamps decide it**, not a version number, with a minute of slack because
  a synced file's stamp can wobble by a second or two. If you ever need to push
  an *older* build, touch the file so its date is newer.
- **Two Excel windows at once.** The second one cannot overwrite a cache the
  first has open. The loader notices and carries on with what is already there.
- **The loader itself** is the one thing that still needs redistributing if it
  changes, which is why it is deliberately small and dull. It should not need to.
