# Handing the add-in to someone else

The `.xlam` is the thing to distribute, not the source. Whoever receives it
needs **no** build step, **no** trust setting and **no** PowerShell policy
change - those are only needed to build it in the first place.

## Once, on your machine

Build it, following [`../addin/INSTALL.md`](../addin/INSTALL.md). You end up
with `dist\WorkbookDoctor.xlam`.

## Then, for everyone else

Put these three files in a folder and zip it:

```
WorkbookDoctor.xlam
Install-WorkbookDoctor.cmd
Uninstall-WorkbookDoctor.cmd
```

They unzip it and double-click the installer. It copies the add-in into their
Excel add-ins folder and switches it on. No admin rights needed.

If Excel will not switch it on by itself the installer says so and gives the
one-time manual tick, which is:

> File → Options → Add-ins → Manage: Excel Add-ins → Go… → tick **Workbook Doctor**

## Why a .cmd and not an .exe

An `.exe` is the format most likely to be quarantined by endpoint security in a
managed environment, and it would need signing to be trusted. A batch file that
copies one file and calls Excel is readable by anyone who wants to check what it
does before running it, which is usually the faster route through IT.

If you do want a single `.exe`, Windows has **IExpress** built in
(`Win+R` → `iexpress`): a wizard that wraps these files into a self-extracting
package that runs the installer. It takes about two minutes and needs nothing
installed. It will still be unsigned, with all that implies.

## A note on updates

Excel holds the `.xlam` open while it is loaded, so an installer cannot
overwrite it in place. The installer will tell you to close Excel if it hits
that. For a real rollout, put the `.xlam` on a network share and point everyone
at that copy instead - then an update is one file replaced, once.
