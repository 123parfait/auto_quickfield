# Versions

## V1 (core features)
1. Connect to QuickField via Python (COM)
2. Modify label values
3. Modify geometry for continuous motion
4. Build mesh and read post-solve data

## V2 (engineered layout)
1. Split functionality into 4 modules for maintainability
2. Simplified duplicate logic (shared open/load helpers)

## Script mapping
- Entry point: `src/main.py` (CLI)
- Connect to QuickField: `src/QF_auto/connection.py`
- Modify label values: `src/QF_auto/labels.py`
- Modify geometry: `src/QF_auto/geometry.py`
- Mesh + read results: `src/QF_auto/solve.py`
