"""Convert the project JSON (see src/core/schema.ts) to pandapower and run a load flow.

Units match the TypeScript schema: kV, kW/kVAr/MVA, km, ohm/km, A.
Motors are modelled as constant-power loads (P = rated kW / efficiency, Q from power factor).
"""
import math

import pandapower as pp


def _cable(project_line, library):
    r, x, amp = project_line.get("rOhmPerKm"), project_line.get("xOhmPerKm"), project_line.get("maxCurrentA")
    ct = project_line.get("cableTypeId")
    if ct is not None:
        found = next((c for c in library.get("cables", []) if c["id"] == ct), None)
        if found is None:
            raise ValueError(f'line {project_line["id"]}: unknown cable type {ct!r}')
        r = found["rOhmPerKm"] if r is None else r
        x = found["xOhmPerKm"] if x is None else x
        amp = found.get("ampacityA") if amp is None else amp
    if r is None or x is None:
        raise ValueError(f'line {project_line["id"]}: missing impedance')
    return r, x, amp


def run_loadflow(project, library=None):
    library = library or {}
    net = pp.create_empty_network(f_hz=project.get("frequencyHz", 50))
    bus = {b["id"]: pp.create_bus(net, vn_kv=b["nominalKv"], name=b["id"]) for b in project.get("buses", [])}

    for s in project.get("sources", []):
        pp.create_ext_grid(net, bus=bus[s["bus"]], vm_pu=s.get("voltagePu", 1.0), name=s["id"])

    trafo, line = {}, {}
    for t in project.get("transformers", []):
        trafo[t["id"]] = pp.create_transformer_from_parameters(
            net, hv_bus=bus[t["hvBus"]], lv_bus=bus[t["lvBus"]], sn_mva=t["ratedMva"],
            vn_hv_kv=t["ratedHvKv"], vn_lv_kv=t["ratedLvKv"], vk_percent=t["vscPercent"],
            vkr_percent=t["vscrPercent"], pfe_kw=t.get("ironLossKw", 0),
            i0_percent=t.get("noLoadCurrentPercent", 0), name=t["id"])
    for l in project.get("lines", []):
        r, x, amp = _cable(l, library)
        line[l["id"]] = pp.create_line_from_parameters(
            net, from_bus=bus[l["fromBus"]], to_bus=bus[l["toBus"]], length_km=l["lengthKm"],
            r_ohm_per_km=r, x_ohm_per_km=x, c_nf_per_km=0.0,
            max_i_ka=(amp / 1000.0) if amp else 1e3, parallel=l.get("parallel", 1), name=l["id"])

    for ld in project.get("loads", []):
        f = ld.get("demandFactor", 1.0) / 1000.0
        pp.create_load(net, bus=bus[ld["bus"]], p_mw=ld["pKw"] * f, q_mvar=ld.get("qKvar", 0.0) * f, name=ld["id"])
    for m in project.get("motors", []):
        p_kw = m["ratedKw"] / m["efficiency"]
        q_kvar = p_kw * math.tan(math.acos(m["powerFactor"]))
        pp.create_load(net, bus=bus[m["bus"]], p_mw=p_kw / 1000.0, q_mvar=q_kvar / 1000.0, name=m["id"])

    for sw in project.get("switches", []):
        closed = sw.get("closed", True)
        if sw["kind"] == "coupler":
            pp.create_switch(net, bus=bus[sw["busA"]], element=bus[sw["busB"]], et="b", closed=closed, name=sw["id"])
        elif sw["elementId"] in line:
            pp.create_switch(net, bus=bus[sw["bus"]], element=line[sw["elementId"]], et="l", closed=closed, name=sw["id"])
        else:
            pp.create_switch(net, bus=bus[sw["bus"]], element=trafo[sw["elementId"]], et="t", closed=closed, name=sw["id"])

    pp.runpp(net, numba=False)

    out_bus = {net.bus.at[i, "name"]: {"vmPu": float(net.res_bus.at[i, "vm_pu"]),
                                       "vaDeg": float(net.res_bus.at[i, "va_degree"])}
               for i in net.bus.index}
    out_line = {net.line.at[i, "name"]: {"loadingPercent": float(net.res_line.at[i, "loading_percent"]),
                                         "currentKa": float(net.res_line.at[i, "i_ka"]),
                                         "lossKw": float(net.res_line.at[i, "pl_mw"]) * 1000.0}
                for i in net.line.index}
    out_trafo = {net.trafo.at[i, "name"]: {"loadingPercent": float(net.res_trafo.at[i, "loading_percent"]),
                                           "lossKw": float(net.res_trafo.at[i, "pl_mw"]) * 1000.0}
                 for i in net.trafo.index}
    return {"converged": bool(net.converged), "buses": out_bus, "lines": out_line, "transformers": out_trafo}
