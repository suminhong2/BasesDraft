# Microsoft Excel version (Office Scripts)

Needs a **work or school Microsoft 365 account** and **Excel for the web** (the Automate tab).
Personal Microsoft accounts don't have Office Scripts.

## Strain collection (`strain-collection/`)
1. Create a blank workbook in Excel for the web, saved in OneDrive or a shared SharePoint / Teams folder.
2. For each file in [`strain-collection/scripts/`](strain-collection/scripts): **Automate → New Script**,
   delete the starter code, paste the file, rename the script to the file name (without `.ts`), **Save script**.
3. Change the lab settings at the top of `SetupWorkbook`, then **Run** it once. It builds the
   Forms, Strains, Requests, Prefixes and Settings tabs. (Or start from
   [`Strain collection.xlsx`](strain-collection/Strain%20collection.xlsx), which is already set up.)
4. On the Settings tab, check the lab code, allele code, next strain number, strain types and the
   emails for request notices.
5. For each script, choose **… → Add in workbook** to put a button on the sheet, and drag it next to
   its section on the Forms tab.

| Script | Forms section | What it does |
|---|---|---|
| `CreateNewStrain` | 1 | Next strain number(s) and allele / transgene names |
| `AddExistingStrain` | 2 | CGC / NBRP strains, or another lab's prefix with the next free number |
| `AddPrefix` | 3 | Registers a new prefix after checking it isn't taken or used |
| `RequestStrain` | 4 | Copies the strain to Requests, checks it isn't out, adds order / WormBase links, @mentions people |

Each section reports in its green **Result** line. If something is missing, the run fails with the
reason. Request emails are sent as @mention comments, which Microsoft 365 emails to the people
mentioned (the workbook must be in OneDrive / SharePoint).

Differences from the Google version: no pop-up forms, no automatic CGC lookup (a link to the CGC page
instead), and no lock against two people adding strains at the same second.
