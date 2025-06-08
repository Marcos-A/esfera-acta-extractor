# Cursor / Windsurf Project Rules

This file defines the core column structure and MP types used throughout the Cursor / Windsurf project. Include this file's content in all AI prompts and code-generation tasks to avoid re-explaining the context.

## Column types

* **RA grade columns**: headers contain the MP code, the RA expression, RA number, etc.
* **CENTRE grade columns**: headers contain the MP code and the literal "CENTRE".
* **EMPRESA grade columns**: headers contain the MP code and the literal "EMPRESA".
* **MP grade columns**: headers contain the MP code alone (no RA, CENTRE, or EMPRESA).

## MP types

* **Type A**: After the RA grade columns, includes three additional MP-related columns:

  1. `MP code + CENTRE` column
  2. `MP code + EMPRESA` column
  3. `MP code` column
* **Type B**: After the RA grade columns, includes only one additional MP-related column:

  * `MP code` column
