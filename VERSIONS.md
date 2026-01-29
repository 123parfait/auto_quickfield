# Versions

## V1 (core features)
1. Connect to QuickField via Python (COM)
2. Modify label values
3. Modify geometry for continuous motion
4. Build mesh and read post-solve data

## Script mapping
- Connect to QuickField: `src/app.py` (`dispatch_qf_app()` / `get_active_problem()`)
- Modify label values: `src/app.py` (`cmd_set_current` / `cmd_label_dump`)
- Modify geometry: `src/app.py` (`cmd_move_block` / `cmd_move_blocks` / `cmd_move_blocks_once` / `cmd_model`)
- Mesh + read results: `src/app.py` (`cmd_solve_integral` / `cmd_solve_force` / `cmd_result_dump`)
