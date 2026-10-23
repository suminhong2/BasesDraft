# MegaBases

Copy-and-paste lab databases for *C. elegans* labs, built on the spreadsheets labs already use.
Each database comes in a **Google Sheets** version (Apps Script) and, where available, a
**Microsoft Excel** version (Office Scripts). No coding needed: you paste the files in and run a setup.

| Database | Google Sheets | Microsoft Excel |
|---|---|---|
| Strain collection | [google-apps-script/strain-collection](google-apps-script/strain-collection) | [microsoft-office-scripts/strain-collection](microsoft-office-scripts/strain-collection) |
| Ordering | [google-apps-script/ordering](google-apps-script/ordering) | not yet |
| Plasmid collection | planned | planned |
| Oligo collection | planned | planned |

## Strain collection
- New strains get the next lab number (e.g. REC1, REC2) and, by strain type, the next allele or
  transgene name (vrt1, vrtSi1, vrtEx1, vrtIs1...).
- Strains from the CGC, NBRP or other labs. Other labs' strains are named with a registered prefix
  and the next free number (OH15903, OH15904...), with checks that a new prefix isn't already taken.
- Requests: copies the strain to a Requests tab, warns if it hasn't been returned yet, links to the
  CGC / NBRP order page and WormBase, and emails chosen people.
- Google version only: CGC lookup that fills in the genotype, pop-up forms, sorting by date added,
  backups to Drive, and sending strain orders to the ordering sheet.

## Ordering
Request form, emails to the people who place orders, an email back to the requester when the order
is placed or arrives, and an Analysis tab with spend per grant, the total, and how often things are ordered.
Strain orders are checked against the strain collection.

See the README in each folder for setup steps.
