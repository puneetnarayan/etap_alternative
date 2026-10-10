import copy
import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent))
from solver.loadflow import run_loadflow  # noqa: E402

SAMPLE = json.loads((pathlib.Path(__file__).resolve().parent.parent / "solver" / "sample.json").read_text())


def test_sample_converges_with_plausible_voltages():
    r = run_loadflow(SAMPLE["project"], SAMPLE["library"])
    assert r["converged"]
    assert r["buses"]["B1"]["vmPu"] == 1.0
    # voltage falls monotonically down the radial feeder, and stays within a sane band
    assert 0.95 < r["buses"]["B3"]["vmPu"] < r["buses"]["B2"]["vmPu"] < 1.0
    assert 0 < r["transformers"]["T1"]["loadingPercent"] < 100


def test_power_balance():
    r = run_loadflow(SAMPLE["project"], SAMPLE["library"])
    p = SAMPLE["project"]
    load_kw = sum(l["pKw"] * l["demandFactor"] for l in p["loads"]) + sum(m["ratedKw"] / m["efficiency"] for m in p["motors"])
    loss_kw = r["lines"]["L1"]["lossKw"] + r["transformers"]["T1"]["lossKw"]
    assert loss_kw > 0 and loss_kw < 0.05 * load_kw


def test_unknown_cable_type_raises():
    bad = copy.deepcopy(SAMPLE)
    bad["project"]["lines"][0]["cableTypeId"] = "nope"
    try:
        run_loadflow(bad["project"], bad["library"])
    except ValueError as e:
        assert "unknown cable type" in str(e)
    else:
        raise AssertionError("expected ValueError")
