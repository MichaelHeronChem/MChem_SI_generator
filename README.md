This script takes my mass and volume data for the 400 imine reaction
experiments in my custom spreadsheet and extracts required data for use in
an SI

Usage: 
1. git clone 
2. UV Sync
3. Make data/raw and data/output
4. Put reaction_planner_tracker.xlsx in data/raw
5. Run uv run python scripts/1_quantity_extractor
6. extracted data is sent to data/output as a csv
7. Run uv run python scripts/2_SI_generator.py
8. SI is sent to data/output as docx.

** LaTeX pdf would be nice