# Migrate-Accordance Version History

## Migrate-Accordance.ps1

* 1.00 - 2026-09-11 - Initial release, Merge and Import not yet confirmed against a real install.
  Inventory, Export, Compare, Merge, and Import modes for the four Windows locations named in the
  vendor migration article: ProgramData modules, LOCALAPPDATA preferences, Documents\Accordance
  Files, and the Oaktree program folder (opt in). Robocopy with file count and byte verification,
  manifest.json, Compare against a staged export with CSV reporting and optional SHA256 content
  checking, additive Merge for module libraries that diverged on both sides, all-or-nothing import
  pre-flight, rename-aside rollback, and -WhatIf support.

  Compare classifies a same-sized file with a differing write time apart from a real content
  difference. Treating the two alike made the first real comparison report 5795 differences out of
  5799 files, which told the operator nothing.
