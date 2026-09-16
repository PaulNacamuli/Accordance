# Accordance

Tooling for moving an Accordance Desktop setup between Windows PCs.

## Migrate-Accordance.ps1

Moves Accordance preferences, user content, and modules from an old PC to a new one.

### Source of the procedure

This script implements the Windows half of the vendor's own instructions:

**<https://support.accordancebible.com/hc/en-us/articles/35804210381851-Migrating-Accordance-Desktop-to-Another-Computer>**

That article names four locations on Windows. The script handles them as four named items, and
nothing beyond them. If the article changes, this script needs revisiting.

| Item | Windows location | Holds |
| --- | --- | --- |
| `Modules` | `C:\ProgramData\Accordance` | modules and support files |
| `Preferences` | `%LOCALAPPDATA%\Accordance` | the `Accordance Preferences` file |
| `UserFiles` | `<Documents>\Accordance Files` | workspaces, highlights, user notes, user tools |
| `Application` | `C:\Program Files (x86)\Oaktree` | the program itself |

The article also says to remove or rename the existing folder on the receiving PC before putting
the old one in place. That is exactly what the import step does, by renaming each target to
`<name>.bak-yyyyMMddHHmmss`.

`AppData` and `ProgramData` are hidden folders in Explorer. The script does not care, but you will
need **View > Show > Hidden items** to inspect them by hand.

### Use it

Close Accordance on both machines first. The script refuses to run while `Accordance.exe` is
alive, because preferences and module indexes are written on exit.

**1. Old PC, see what is there.** Copies nothing.

```powershell
.\Migrate-Accordance.ps1 -Mode Inventory
```

**2. Old PC, stage the data** onto a USB drive, external disk, or share both PCs can reach. Use
the size that Inventory reported to pick a drive.

```powershell
.\Migrate-Accordance.ps1 -Mode Export -Path E:\
```

This writes `E:\AccordanceMigration` containing one folder per item, a `manifest.json`, and a log.

**3. New PC, install Accordance normally.** Launch it once so it creates its own folders, then
close it.

**4. New PC, compare before you overwrite anything.** Changes nothing.

```powershell
.\Migrate-Accordance.ps1 -Mode Compare -Path E:\ -ReportPath C:\Temp\AccordanceCompare -Hash
```

Per item this reports how many files are identical, how many genuinely differ, how many exist only
in the export, and how many exist only on the new PC. That last number decides everything: those
files are what an Import would move into the `.bak` folder. The CSVs list every one of them.

`-Hash` is worth the extra minutes. Without it, files that match in size but not in write time are
counted separately and left undecided, because a timestamp difference on its own proves nothing.
Two machines that installed the same module at different times produce exactly that, by the
thousand. `-Hash` settles those cases by comparing content.

**5. New PC, import what the old PC owns outright.** Run PowerShell **as administrator**.

```powershell
.\Migrate-Accordance.ps1 -Mode Import -Path E:\ -Force -SkipItem Modules
```

Import replaces a whole folder, so use it where the old PC's version is the one you want
throughout: preferences and user content. Check Compare first and confirm "only on this PC" is zero
for those items, which means nothing gets buried. Add `-WhatIf` to preview.

**6. New PC, merge the modules instead of replacing them.** Still elevated.

```powershell
.\Migrate-Accordance.ps1 -Mode Merge -Path E:\ -SkipItem Preferences,UserFiles
```

Merge only adds files this PC does not have. Nothing existing is renamed, replaced, or deleted. Use
it wherever Compare shows files on both sides that the other lacks, which is the normal state of
two module libraries that each downloaded independently.

**7. Start Accordance** and check your workspaces, user notes, highlights, and module list. Once
satisfied, delete the `*.bak-*` folders the import left behind. Merge creates none.

### Parameters

| Parameter | Effect |
| --- | --- |
| `-Mode` | `Inventory`, `Export`, `Compare`, `Merge`, or `Import`. Required. |
| `-Path` | Staging folder. Required for every mode except Inventory. The same value works throughout. |
| `-ReportPath` | Compare only. Folder for the per-item CSVs. Without it, Compare prints counts and a sample and writes nothing. |
| `-Hash` | Compare only. SHA256 the files that match in size but not write time, to decide them for certain. |
| `-Force` | Import: allows an existing target folder to be renamed aside. Export: allows writing into a staging folder that already holds an export. |
| `-IncludeApplication` | Also migrate `C:\Program Files (x86)\Oaktree`. Off by default. |
| `-SkipItem` | Leave named items alone, e.g. `-SkipItem Modules,UserFiles`. |
| `-WhatIf` | Show everything that would happen, change nothing, write no log. |

### Design notes

**Neither module library is likely to be a superset of the other.** Accordance downloads modules
against your account, so each machine's `C:\ProgramData\Accordance` reflects whatever it happened to
download. In practice the old PC holds a module or two the new one never fetched, and the new one
holds several the old one never had. Import would gain the first set at the cost of burying the
second; skipping Modules does the reverse. Merge is the way out, because it only adds. This is the
reason Merge exists.

**Import replaces, Merge adds.** That is the whole distinction. Import renames the target folder
aside and puts the old PC's version in its place, which is what you want for preferences and user
content, where the old machine's state is authoritative. Merge copies in only the files that are
absent here and leaves everything else untouched, which is what you want for a module library.

**Timestamps are not evidence.** Two machines that installed the same module at different times
produce two identical files with different write times. Compare counts those separately rather than
calling them differences, and `-Hash` resolves them by content. Only a size difference, or a hash
difference, means the content actually differs.

**The program folder is excluded by default.** A fresh install on the new PC already supplies the
program files, usually a newer build than the old PC had. Copying the old ones over it is a
downgrade, not a migration. `-IncludeApplication` is there for the case where you want the old
binaries anyway.

**Import is all or nothing.** Every check runs before the first rename or copy: manifest present,
staged data matching the manifest byte for byte, elevation available, targets safe to move aside.
If any item fails a check, nothing is touched at all. A partial import is the one outcome worth
avoiding, because new modules against old preferences is a state neither PC ever had, and nothing
inside Accordance makes it obvious that it happened. Name the blocked items in `-SkipItem` if you
want the rest to proceed anyway.

**Paths are resolved live on each machine,** never read back from the manifest. The two PCs can
have different user names, and `Documents` can be redirected into OneDrive on either side or
neither. On import, an `Accordance Files` folder that already exists wins, since that is where the
new PC's Accordance just created it.

**Copying is robocopy,** for predictable handling of multi-gigabyte module libraries and long
paths. Attributes and timestamps come across; ACLs deliberately do not, since they would refer to
the old PC's account. Every item is verified afterwards by comparing file count and total bytes
against the source, and a mismatch is reported as a failure.

**Rollback** is the `<name>.bak-yyyyMMddHHmmss` folders. Nothing is deleted, so a bad import is
undone by deleting the new folder and renaming the backup back.

### What this does not cover

- **macOS.** The article covers Mac paths too; this script is Windows only.
- **Licensing and authorization.** The article says nothing about deauthorizing the old machine or
  authorizing the new one. If Accordance asks you to sign in or authorize after the migration,
  that is outside what this script touches.
- **Installing Accordance.** Step 3 is a normal manual install.

### Version history

See [Migrate-Accordance.ver.md](Migrate-Accordance.ver.md).

### Testing status

Exercised end to end against a sandboxed folder tree: 67 assertions covering export, manifest
accuracy, verification, every Compare classification including the timestamp-only case and `-Hash`,
Merge's additive behaviour and its no-op on a second run, the rename-aside, restore onto both
populated and empty targets, the all-or-nothing abort, `-WhatIf` on each writing mode, tampered
staging, and every refusal path.

Inventory, Export, and Compare have been run for real between two PCs. Merge and Import have **not**
yet been run against a real Accordance install.
