# ARWC-Pilot

Website of the Asian Registry of Wrongful Convictions. The whole site is `index.html`.

## Turning "Cases under review" on or off

Near the top of the script in `index.html`, under `SITE`, change

    showInventory: true,

to `false` to hide the page and its menu links, or back to `true` to show them again. On GitHub,
edit the file in the browser and commit to `main`; the site updates within a few minutes.

Hiding the page does not remove its data from `index.html`: anyone reading the page source can
still see it. To take the list down completely, remove the data as well (see below), or ask for
the page to be removed.

## Updating the list

The list comes from the "Cases" sheet of the case inventory workbook, which is kept outside this
repository. After editing the workbook, run

    python3 tools/build_inventory_html.py ARWC_Case_Inventory.xlsx --site index.html

and commit `index.html`. Registry records are left out, as are the internal columns (preliminary
category, finality guess, internal notes, screening codes). Screened-out rows show a standard
public reason taken from their exclusion code in column U (X-01 to X-04), not the internal note.
Rows coded X-05 (no official correction yet) are left off the public page.
Where a row has an official document (columns Q-S), its link is shown as the source. To publish an
empty list, put a workbook with no candidate rows through the same command.
