# Physician Name to NPI Generator

## Method

1. Use https://npiregistry.cms.hhs.gov/
2. Enter first and last name
3. Run search with "check this box to search for Exact Matches"

**Output**

1. address
2. primary taxonomy
3. NPI

## Edge Cases

1. If a providers name does not show?

- Sometimes they move states, so they do not show up

2. What if there are multiple providers of the same name?

- Show them all

3. What if the provider has multiple last names

- For now leave them blank, and fill them in manually later
