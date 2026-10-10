# Spike: pandapower on Vercel

Status: **local verification done; Vercel deployment not yet tested.**

## Verified locally (Python 3.13, pandapower 3.5.6, numpy 2.4.6, scipy 1.18.1, pandas 2.3.3)
- `solver/loadflow.py` converts the project JSON to pandapower and runs a load flow; `api/loadflow.py` exposes it as `POST /api/loadflow`.
- Sample plant converges (B2 ≈ 0.992 pu, B3 ≈ 0.983 pu; T1 ≈ 24 % loaded); `python -m pytest tests_py`.
- Installed footprint: ~378 MB as installed by pip; ~263 MB after removing pip/pytest, `__pycache__` and test directories.
  pandapower still imports `pandapower.networks` (~26 MB), so that cannot be dropped without patching.
- Cold-start import + solve measured only on a dev machine (~0.3 s solve, import time not separately measured).

## Not verified
- Vercel's current Python bundle-size limit (the docs confirm a limit exists and recommend `excludeFiles`; I could not retrieve the number).
- That a Next.js app plus `api/*.py` deploy together from one repo.
- Real cold-start latency on Vercel.
- pandapower without numba prints a warning and may be slower on large networks.

## Fallbacks if the bundle does not fit
1. Pin/slim dependencies further (e.g. drop unused packages).
2. Host the Python solver as a separate service (free tiers exist; verify terms) and call it from the Next.js app.
3. Replace pandapower with a small own Newton-Raphson solver in TypeScript (more work, full control).
