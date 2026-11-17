# Grace-Period Client Contact Extractor

Every month the insurer sends a list of policies in their **grace period** (payment overdue). This tool enriches that list with the contact details needed to reach each client before the policy is canceled.

## What it does (`FullLogict.py`)
1. Normalizes policy numbers (spaces, dashes, leading zeros) so they match across systems
2. Removes duplicate policies and splits the list into unique vs. duplicated records
3. Splits the insured's full name into first/last name (handles `"Last, First"` and 2-, 3- and 4-word names)
4. Looks up each client's **email** in the full client base and validates it
5. Adds the assigned **sales advisor's email and alias**, so each advisor can follow up with their own clients
6. Exports a ready-to-use Excel file

`Main.py`, `MatchExels.py` and `FNames&LNames.py` are the earlier single-step versions of the same process.

## Tech stack
Python · pandas · regex

## Run it
```bash
pip install -r requirements.txt
python FullLogict.py
```

> **Note:** The input files contain confidential client and company data, so they are not included in this repository. Only the source code is published.
