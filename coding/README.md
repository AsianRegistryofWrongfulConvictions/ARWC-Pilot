# Independent coding

Second-coder checks of the coded cases, to test whether Codebook v1.0 leads two coders to the
same coding and to find where it does not.

## Round 1

Cases: ARWC-TWN-001 (Chiang Kuo-ching), ARWC-CHN-002 (Zhang Yuhuan), ARWC-CHN-001 (Nie Shubin),
ARWC-SGP-001 (Ilechukwu) and ARWC-IND-001 (Ankush Maruti Shinde and five others).

1. Give the coder `ARWC_Coding_Round1.xlsx`, the framework paper (Supplement A is the codebook)
   and the link to the site's Methodology and Glossary pages. The template holds the case IDs,
   the names of the people convicted and the sources, and none of the reference coding.
2. Ask them not to open the site's record pages for these cases, the coded-set documents or
   `data/ARWC_Case_Inventory.xlsx` until they have finished.
3. When the file comes back, save it as `ARWC_Coding_Round1_<initials>.xlsx` and run:

   ```
   python3 tools/compare_coding.py coding/ARWC_Coding_Round1_<initials>.xlsx data/ARWC_Case_Inventory.xlsx coding/Agreement_Round1_<initials>.xlsx
   ```

   Open the report in Excel (formulas calculate on opening). It has field-by-field results,
   agreement by case, Cohen's kappa for the checkbox items (categories, factors and triggers),
   and the coder's time log and notes on where the codebook was unclear.
4. Go through each disagreement together. Record the outcome in the "Discussion / decision"
   column, and add new rules to the decision log.

Note: the reference coding for ARWC-CHN-001 is the Supplement B entry, coded at v0.3.

To make a packet for other cases or another round:

```
python3 tools/make_coding_packet.py data/ARWC_Case_Inventory.xlsx coding/ARWC_Coding_Round2.xlsx <case IDs>
```
