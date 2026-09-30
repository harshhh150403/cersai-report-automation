"""CERSAI SI registration batch tool.

Three steps, run one after another:
  1. Build the working Excel and the required-details file from CKYC, COLLATERAL
     and the legal title-report PDFs.
  2. Apply the values typed into the required-details file and build the final
     Excel for review.
  3. Check the reviewed final Excel and write the .DAT file for CERSAI.

Start it with the .bat files in the project folder, or with
    python -m cersai_tool run1 | run2 | run3
"""
