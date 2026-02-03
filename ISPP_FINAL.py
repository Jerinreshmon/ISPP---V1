#ISPP FINAL.py
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, scrolledtext
import pandas as pd
from datetime import datetime
import os

# ------------------ CONSTANTS ------------------
standard_copies = [1e5, 1e4, 2e3, 2e2, 2e1, 1e1, 5, 2.5, 1.25, 0.625]
copy_options_str = [str(x) for x in standard_copies]
targets = ["FluA", "FluB", "RSV A", "RSV B", "SC2"]
dyes = [
    "FAM", "HEX", "TET", "DY 480", "ROX", "Texas Red",
    "Atto 565", "AF 568", "Atto 633", "Alexa Fluor 647",
    "Atto 660", "AF 700"
]
excel_file = "Sample_Info-Sheet.xlsx"


# ------------------ FUNCTIONS ------------------
def agent_preparation(desired_copies, A_uL, dilution_type):
    """
    Agent Prep Guide:
      - Stock = 1e6 copies/µL (virus stock)
      - Desired_copies = copies you want PER REACTION
      - A_uL = Agent volume per reaction (µL)
      - We compute target concentration C_target = desired_copies / A (copies/µL)
      - Build a stepwise dilution series from 1e6 down to ~C_target
      - All steps use 100 µL final volume with stock and water in multiples of 10 µL.
    """
    output = []
    numeric = {}

    # Parse A
    try:
        A = float(A_uL)
    except Exception:
        raise ValueError("Agent volume A must be numeric.")
    if A <= 0:
        raise ValueError("Agent volume A must be positive.")

    # Parse desired copies
    try:
        desired = float(desired_copies)
    except Exception:
        raise ValueError("Desired copies must be numeric.")
    if desired <= 0:
        raise ValueError("Desired copies must be positive.")

    stock = 1e6  # copies/µL
    final_vol = 100.0  # µL for each dilution step

    # Target concentration in Agent tube so that A µL contains 'desired' copies
    C_target = desired / A  # copies/µL

    if C_target >= stock:
        raise ValueError(
            f"Target concentration ({C_target:.3g} copies/µL) cannot be higher than stock (1e6 copies/µL). "
            f"Increase A or lower desired copies."
        )

    output.append("===== AGENT PREP GUIDE (USING AGENT VOLUME A) =====")
    output.append("")
    output.append(f"Stock concentration: 1.0e6 copies/µL")
    output.append(f"Agent volume A per reaction: {A:.3f} µL")
    output.append(f"Desired copies per reaction: {desired:.3g} copies")
    output.append(f"→ Target concentration in Agent tube: C_target = desired / A = {C_target:.6g} copies/µL")
    output.append("")
    output.append("All dilution steps below use 100 µL final volume and volumes in multiples of 10 µL.")
    output.append("")

    steps = []
    current = stock

    if dilution_type == "Direct":
        # Direct single-step dilution from 1e6 → C_target (approx) with multiples of 10 µL
        desired_factor = C_target / current  # C2 = C1 * (V_stock/V_total)
        raw_V_stock = final_vol * desired_factor
        # Round stock volume to nearest 10 µL (between 10 and 90)
        V_stock = max(10, min(90, round(raw_V_stock / 10) * 10))
        V_water = final_vol - V_stock
        actual_factor = V_stock / final_vol
        C_final = current * actual_factor
        copies_in_A = C_final * A

        output.append("DIRECT DILUTION (approximate, with 10 µL increments):")
        output.append(
            f"  From 1.0e6 copies/µL → target ~{C_target:.6g} copies/µL in one step (100 µL total):"
        )
        output.append(
            f"    Mix {V_stock:.0f} µL of 1.0e6 stock + {V_water:.0f} µL water → 100 µL of Agent Mix"
        )
        output.append("")
        output.append(f"  Resulting concentration: {C_final:.6g} copies/µL")
        output.append(
            f"  Copies in A = {A:.3f} µL × {C_final:.6g} copies/µL = {copies_in_A:.3g} copies"
        )
        output.append(
            f"  (Target was {desired:.3g} copies per {A:.3f} µL; difference = {copies_in_A - desired:.3g} copies)"
        )

        numeric.update({
            "A_uL": A,
            "desired_copies": desired,
            "C_target_copies_per_uL": C_target,
            "V_stock_direct_uL": V_stock,
            "V_water_direct_uL": V_water,
            "C_final_direct_copies_per_uL": C_final,
            "copies_in_A_direct": copies_in_A,
        })

        return "\n".join(output), numeric

    # STEPWISE MODE
    output.append("STEPWISE DILUTION SERIES:")
    output.append("")

    # 1) Do as many clean 10× dilutions as possible using 10 µL + 90 µL.
    while current / 10 >= C_target:
        next_conc = current / 10.0
        steps.append({
            "from": current,
            "to": next_conc,
            "V_stock": 10.0,
            "V_water": 90.0
        })
        current = next_conc

    # 2) If we still haven't reached C_target exactly, do one final approximate step
    if current > C_target:
        desired_factor = C_target / current
        raw_V_stock = final_vol * desired_factor
        V_stock = max(10, min(90, round(raw_V_stock / 10) * 10))
        V_water = final_vol - V_stock
        actual_factor = V_stock / final_vol
        C_final = current * actual_factor
        steps.append({
            "from": current,
            "to": C_final,
            "V_stock": V_stock,
            "V_water": V_water
        })
        current = C_final
    else:
        C_final = current

    # 3) Report series
    for i, step in enumerate(steps, start=1):
        output.append(
            f"Step {i}: {step['from']:.3g} → {step['to']:.3g} copies/µL"
        )
        output.append(
            f"    Mix {step['V_stock']:.0f} µL of {step['from']:.3g} stock + "
            f"{step['V_water']:.0f} µL water → 100 µL at {step['to']:.3g} copies/µL"
        )
        output.append("")

    copies_in_A = C_final * A

    output.append("FINAL AGENT TUBE:")
    output.append(f"  Final concentration ~ {C_final:.6g} copies/µL")
    output.append(
        f"  Copies in A = {A:.3f} µL × {C_final:.6g} copies/µL = {copies_in_A:.3g} copies"
    )
    output.append(
        f"  Target was {desired:.3g} copies per {A:.3f} µL; difference = {copies_in_A - desired:.3g} copies"
    )

    numeric.update({
        "A_uL": A,
        "desired_copies": desired,
        "C_target_copies_per_uL": C_target,
        "final_copies_per_uL": C_final,
        "copies_in_A_final": copies_in_A,
        "num_steps": len(steps),
    })
    for idx, step in enumerate(steps, start=1):
        numeric[f"step{idx}_from"] = step["from"]
        numeric[f"step{idx}_to"] = step["to"]
        numeric[f"step{idx}_V_stock_uL"] = step["V_stock"]
        numeric[f"step{idx}_V_water_uL"] = step["V_water"]

    return "\n".join(output), numeric


# ---------- FINAL SAMPLE PREP (single-plex + multiplex) ----------
def final_sample_preparation_new(R_uL, A_uL, O_uL, X_uM, target_choice, dye_choice, N_samples):
    """
    SINGLE-PLEX Final Sample Preparation logic with N samples.
    (unchanged from previous version)
    """
    output = []
    numeric = {}

    # Parse numeric inputs
    try:
        R = float(R_uL)
        A = float(A_uL)
        O = float(O_uL)
        X = float(X_uM)
    except Exception:
        raise ValueError("R, A, O, X must be numeric.")

    if R <= 0 or A < 0:
        raise ValueError("R must be positive and A must be non-negative.")

    # Parse N (number of samples)
    if N_samples is None or str(N_samples).strip() == "":
        N = 1
    else:
        try:
            N = int(str(N_samples).strip())
        except Exception:
            raise ValueError("Number of samples N must be an integer.")

    if N < 1 or N > 15:
        raise ValueError("Number of samples N must be between 1 and 15.")

    # Determine multiplier m for master mixes based on N
    if N == 1:
        m = 1
    elif N in (2, 3):
        m = 5
    elif N in (4, 5, 6):
        m = 10
    elif N in (7, 8, 9):
        m = 15
    elif N in (10, 11, 12):
        m = 20
    else:  # 13, 14, 15
        m = 25

    output.append(f"Number of samples requested: N = {N}.")
    output.append(f"For master mixes, all per-reaction volumes of R, A, and R-components will be multiplied by {m}×.")
    output.append("")

    # Per reaction final volume
    F = R + A
    output.append(f"Computed final mix volume per reaction: F = R + A = {R:.3f} + {A:.3f} = {F:.3f} µL")

    # H2O inside Reagent R
    H20 = R - (10.0 + 0.8 + O)
    if H20 < 0:
        raise ValueError("Computed water in Reagent (H₂O) is negative. Check R and O values.")

    # Step 1: J for ONE reaction
    J = ((X * F) / 100.0) * 12.5
    output.append("")
    output.append("Step 1 — Calculate J (Volume of Probe that you add to the Probe-Primer Mix for ONE reaction):")
    output.append("                                                         ")
    output.append(
        f"  J = ((X × F) / 100) × 12.5 = (({X:.4f} µM × {F:.4f} µL)/100) × 12.5 = {J:.6f} µL"
    )

    # Step 2: Base 25 µL Probe-Primer Mix
    mix_total_base = 25.0
    remaining = mix_total_base - J
    if remaining <= 0:
        raise ValueError(
            "J is too large: not enough volume left for forward/reverse primers and water in the 25 µL base mix."
        )

    # Primer volumes for base mix
    preferred_primer = 2.0
    max_total_primers = 2 * preferred_primer
    if remaining >= max_total_primers:
        fwd = preferred_primer
        rev = preferred_primer
    else:
        fwd = rev = remaining / 2.0
        if fwd > 2.0:
            fwd = rev = 2.0

    water_mix = mix_total_base - (J + fwd + rev)

    probe_name = f"{dye_choice}-{target_choice}"
    fwd_name = f"{target_choice} Forward Primer"
    rev_name = f"{target_choice} Reverse Primer"

    output.append("                            ")
    output.append("Step 2 — Prepare a 25 µL Probe-Primer Mix (base composition):")
    output.append("                                                ")
    output.append(f"  • \"{probe_name}\" Probe (from J): {J:.6f} µL")
    output.append(f"  • Forward primer: {fwd:.6f} µL")
    output.append(f"  • Reverse primer: {rev:.6f} µL")
    output.append(f"  • H₂O (to reach 25 µL): {water_mix:.6f} µL")
    output.append("                                                 ")
    output.append(
        f"  → To Conclude (base): Add {J:.6f} µL of \"{probe_name}\" Probe with {fwd:.6f} µL of \"{fwd_name}\" and "
        f"{rev:.6f} µL of \"{rev_name}\" with {water_mix:.6f} µL H₂O to make a 25 µL Probe-Primer Mix."
    )

    # N-dependent larger Probe-Primer Mix for ALL reactions
    if N == 1:
        mix_total_scaled = mix_total_base
    elif N in (2, 3):
        mix_total_scaled = 50.0
    elif N in (4, 5, 6, 7, 8):
        mix_total_scaled = 100.0
    elif N in (9, 10, 11, 12):
        mix_total_scaled = 150.0
    else:  # 13, 14, 15
        mix_total_scaled = 200.0

    scale_factor = mix_total_scaled / mix_total_base
    J_scaled = J * scale_factor
    fwd_scaled = fwd * scale_factor
    rev_scaled = rev * scale_factor
    water_scaled = water_mix * scale_factor

    if N > 1:
        output.append("")
        output.append(
            f"For multiple samples (N = {N}), prepare a larger Probe-Primer Mix so you have enough for ALL reactions:"
        )
        output.append(f"  Recommended total Probe-Primer Mix volume: {mix_total_scaled:.1f} µL")
        output.append(f"  Multiply each 25 µL-base component by {scale_factor:.2f} so they sum to {mix_total_scaled:.1f} µL:")
        output.append(f"    • \"{probe_name}\" Probe: {J_scaled:.6f} µL")
        output.append(f"    • Forward primer: {fwd_scaled:.6f} µL")
        output.append(f"    • Reverse primer: {rev_scaled:.6f} µL")
        output.append(f"    • H₂O: {water_scaled:.6f} µL")
        output.append(
            f"    → This {mix_total_scaled:.1f} µL Probe-Primer Mix will be the source of oligos for your N = {N} reactions."
        )

    output.append("")
    output.append("The following steps (3–5) describe how to prepare ONE reaction (single sample).")

    # Step 3: Reagent Mix (per reaction)
    output.append("                                                      ")
    output.append("Step 3 — Reagent Mix Preparation (per reaction)")
    output.append(
        f"  • Prepare your Reagent R of {R:.3f} µL with the following components per reaction:"
    )
    output.append(f"      - 10.0 µL PS2XII")
    output.append(f"      - 0.8 µL AuNPs")
    output.append(f"      - {O:.3f} µL oligos (taken from the Probe-Primer Mix)")
    output.append(f"      - {H20:.3f} µL H₂O")
    output.append(f"    → Total per reaction Reagent volume: {R:.3f} µL")
    output.append("                                ")
    output.append(
        f"  • To make this, first add PS2XII and the AuNPs, then pipette {O:.3f} µL from the Probe-Primer Mix into "
        f"the Reagent Mix, and finally add {H20:.3f} µL H₂O."
    )

    # Step 4: Agent Mix Preparation (per reaction)
    output.append("                                 ")
    output.append("Step 4 — Agent Mix Preparation (per reaction):")
    output.append("                                           ")
    output.append(
        f"  • Prepare the Agent mix ({A:.3f} µL per reaction) as calculated and instructed by the Agent Prep Guide."
    )

    # Step 5: Final Mix (per reaction)
    output.append("Step 5 — Preparing the Final Mix (per reaction)")
    output.append(
        f"  • For ONE reaction, pipette {R:.3f} µL Reagent and {A:.3f} µL Agent into the final reaction tube "
        f"to make {F:.3f} µL final mix."
    )
    output.append("                         ")

    # Step 6: Master mixes for N samples using multiplier m
    R_total = R * m
    A_total = A * m
    F_total = F * m

    PS2XII_total = 10.0 * m
    AuNPs_total = 0.8 * m
    O_total = O * m
    H20_total = H20 * m

    output.append("")
    output.append("Step 6 — Preparing Master Mixes for N Reactions")
    output.append("                                                ")
    output.append(
        f"  • You are planning to run N = {N} reactions. To account for pipetting loss, "
        f"prepare {m}× the per-reaction volume of each mix (R, A, and R-components)."
    )
    output.append("")
    output.append("  Reagent R MASTER MIX (for all reactions):")
    output.append(f"    - PS2XII: 10.0 µL × {m} = {PS2XII_total:.3f} µL")
    output.append(f"    - AuNPs: 0.8 µL × {m} = {AuNPs_total:.3f} µL")
    output.append(f"    - Oligos from Probe-Primer Mix: {O:.3f} µL × {m} = {O_total:.3f} µL")
    output.append(f"    - H₂O: {H20:.3f} µL × {m} = {H20_total:.3f} µL")
    output.append(f"    → Total Reagent R master mix volume: {R_total:.3f} µL")
    output.append("")
    output.append("  Agent MASTER MIX:")
    output.append(f"    - Agent volume per reaction: {A:.3f} µL")
    output.append(f"    - Prepare: {A:.3f} µL × {m} = {A_total:.3f} µL of Agent mix in total.")
    output.append("")
    output.append("  Final Mix TOTAL VOLUME prepared:")
    output.append(f"    - Total Master Mix Prepared: {F_total:.3f} µL")
    output.append(
        f"    - You prepare this by combining the Reagent Master Mix volume ({R_total:.3f} µL) "
        f"and the Agent Master Mix volume ({A_total:.3f} µL)."
    )
    output.append("")
    output.append(
        f"  → From these master mixes, dispense {R:.3f} µL of Reagent and {A:.3f} µL of Agent into each tube "
        f"to set up N = {N} reactions."
    )
    output.append("")
    output.append("YOU ARE GOOD TO GO :) !")

    numeric = {
        "R_uL": float(R),
        "A_uL": float(A),
        "F_uL": float(F),
        "O_uL": float(O),
        "X_uM": float(X),
        "H2O_in_R_uL": float(H20),
        "J_uL": float(J),
        "mix_total_base_uL": float(mix_total_base),
        "mix_forward_uL": float(fwd),
        "mix_reverse_uL": float(rev),
        "mix_h2o_uL": float(water_mix),
        "N_samples": int(N),
        "multiplier": float(m),
        "probe_mix_total_scaled_uL": float(mix_total_scaled),
        "probe_J_scaled_uL": float(J_scaled),
        "probe_fwd_scaled_uL": float(fwd_scaled),
        "probe_rev_scaled_uL": float(rev_scaled),
        "probe_h2o_scaled_uL": float(water_scaled),
        "R_total_uL": float(R_total),
        "A_total_uL": float(A_total),
        "F_total_uL": float(F_total),
        "PS2XII_total_uL": float(PS2XII_total),
        "AuNPs_total_uL": float(AuNPs_total),
        "O_total_uL": float(O_total),
        "H2O_total_uL": float(H20_total),
    }

    output_text = "\n".join(output)
    return output_text, numeric


def final_sample_preparation_multiplex(R_uL, A_uL, O_uL, X_list, target_list, dye_list, N_samples):
    """
    MULTIPLEX Final Sample Preparation logic with N samples and L probe–target pairs.
    (unchanged from previous version)
    """
    output = []
    numeric = {}

    # Parse numeric R, A, O
    try:
        R = float(R_uL)
        A = float(A_uL)
        O = float(O_uL)
    except Exception:
        raise ValueError("R, A, O must be numeric.")

    if R <= 0 or A < 0:
        raise ValueError("R must be positive and A must be non-negative.")

    # Parse N
    if N_samples is None or str(N_samples).strip() == "":
        N = 1
    else:
        try:
            N = int(str(N_samples).strip())
        except Exception:
            raise ValueError("Number of samples N must be an integer.")

    if N < 1 or N > 15:
        raise ValueError("Number of samples N must be between 1 and 15.")

    # Determine multiplier m for master mixes based on N
    if N == 1:
        m = 1
    elif N in (2, 3):
        m = 5
    elif N in (4, 5, 6):
        m = 10
    elif N in (7, 8, 9):
        m = 15
    elif N in (10, 11, 12):
        m = 20
    else:  # 13, 14, 15
        m = 25

    # Number of plexes
    L = len(X_list)
    if L < 1 or L > 7:
        raise ValueError("Number of plexes L must be between 1 and 7.")

    # Parse X values
    X_vals = []
    for i, x in enumerate(X_list):
        try:
            X_vals.append(float(str(x).strip()))
        except Exception:
            raise ValueError(f"X for pair {i+1} must be numeric.")

    # Per reaction final volume
    F = R + A
    output.append(f"Number of samples requested: N = {N}.")
    output.append(f"Number of plexes (L): {L}.")
    output.append(f"For master mixes, all per-reaction volumes of R, A, and R-components will be multiplied by {m}×.")
    output.append("")
    output.append(f"Computed final mix volume per reaction: F = R + A = {R:.3f} + {A:.3f} = {F:.3f} µL")

    # H2O inside Reagent R
    H20 = R - (10.0 + 0.8 + O)
    if H20 < 0:
        raise ValueError("Computed water in Reagent (H₂O) is negative. Check R and O values.")

    # Step 1 — J values for each probe
    output.append("")
    output.append("Step 1 — Calculate J for EACH probe (volume of each probe added to the Probe-Primer Mix for ONE reaction):")
    output.append("")

    J_list = []
    for i in range(L):
        X_i = X_vals[i]
        tgt_i = target_list[i]
        dye_i = dye_list[i]
        J_i = ((X_i * F) / 100.0) * 12.5
        J_list.append(J_i)
        output.append(
            f"  Pair {i+1}: {dye_i}-{tgt_i}: "
            f"J_{i+1} = (({X_i:.4f} µM × {F:.4f} µL)/100) × 12.5 = {J_i:.6f} µL"
        )

    # Step 2 — Base 25 µL multiplex Probe-Primer Mix
    mix_total_base = 25.0
    sum_J = sum(J_list)
    primers_per_pair_fwd = 2.0
    primers_per_pair_rev = 2.0
    total_primers = L * (primers_per_pair_fwd + primers_per_pair_rev)
    water_mix = mix_total_base - (sum_J + total_primers)

    if water_mix <= 0:
        raise ValueError(
            "Probe + primer volumes exceed 25 µL. Reduce X values or primer volumes."
        )

    output.append("")
    output.append("Step 2 — Prepare a 25 µL MULTIPLEX Probe-Primer Mix (base composition):")
    output.append("")

    # Probes
    for i in range(L):
        tgt_i = target_list[i]
        dye_i = dye_list[i]
        J_i = J_list[i]
        output.append(f"  • \"{dye_i}-{tgt_i}\" Probe (from J_{i+1}): {J_i:.6f} µL")

    # Primers
    for i in range(L):
        tgt_i = target_list[i]
        output.append(f"  • {tgt_i} Forward primer: {primers_per_pair_fwd:.6f} µL")
        output.append(f"  • {tgt_i} Reverse primer: {primers_per_pair_rev:.6f} µL")

    output.append(f"  • H₂O (to reach 25 µL): {water_mix:.6f} µL")
    output.append("")
    output.append(
        "  → To Conclude (base): Combine all listed probe and primer volumes above with the H₂O volume "
        "to make a 25 µL MULTIPLEX Probe-Primer Mix."
    )

    # N-dependent larger Probe-Primer Mix
    if N == 1:
        mix_total_scaled = mix_total_base
    elif N in (2, 3):
        mix_total_scaled = 50.0
    elif N in (4, 5, 6, 7, 8):
        mix_total_scaled = 100.0
    elif N in (9, 10, 11, 12):
        mix_total_scaled = 150.0
    else:  # 13, 14, 15
        mix_total_scaled = 200.0

    scale_factor = mix_total_scaled / mix_total_base
    J_scaled_list = [J_i * scale_factor for J_i in J_list]
    primers_fwd_scaled = primers_per_pair_fwd * scale_factor
    primers_rev_scaled = primers_per_pair_rev * scale_factor
    water_scaled = water_mix * scale_factor

    if N > 1:
        output.append("")
        output.append(
            f"For multiple samples (N = {N}), prepare a larger MULTIPLEX Probe-Primer Mix so you have enough for ALL reactions:"
        )
        output.append(f"  Recommended total Probe-Primer Mix volume: {mix_total_scaled:.1f} µL")
        output.append(
            f"  Multiply each 25 µL-base component by {scale_factor:.2f} so they sum to {mix_total_scaled:.1f} µL:"
        )
        output.append("")
        for i in range(L):
            tgt_i = target_list[i]
            dye_i = dye_list[i]
            output.append(
                f"    • \"{dye_i}-{tgt_i}\" Probe: {J_scaled_list[i]:.6f} µL"
            )
        for i in range(L):
            tgt_i = target_list[i]
            output.append(f"    • {tgt_i} Forward primer: {primers_fwd_scaled:.6f} µL")
            output.append(f"    • {tgt_i} Reverse primer: {primers_rev_scaled:.6f} µL")
        output.append(f"    • H₂O: {water_scaled:.6f} µL")
        output.append(
            f"    → This {mix_total_scaled:.1f} µL MULTIPLEX Probe-Primer Mix will be the source of oligos for your N = {N} reactions."
        )

    output.append("")
    output.append("The following steps (3–5) describe how to prepare ONE multiplex reaction (single sample).")

    # Step 3: Reagent Mix (per reaction)
    output.append("")
    output.append("Step 3 — Reagent Mix Preparation (per reaction)")
    output.append(
        f"  • Prepare your Reagent R of {R:.3f} µL with the following components per reaction:"
    )
    output.append(f"      - 10.0 µL PS2XII")
    output.append(f"      - 0.8 µL AuNPs")
    output.append(f"      - {O:.3f} µL MULTIPLEX oligos (taken from the Probe-Primer Mix)")
    output.append(f"      - {H20:.3f} µL H₂O")
    output.append(f"    → Total per reaction Reagent volume: {R:.3f} µL")
    output.append("")
    output.append(
        f"  • To make this, first add PS2XII and the AuNPs, then pipette {O:.3f} µL from the MULTIPLEX Probe-Primer Mix "
        f"into the Reagent Mix, and finally add {H20:.3f} µL H₂O."
    )

    # Step 4: Agent Mix Preparation (per reaction)
    output.append("")
    output.append("Step 4 — Agent Mix Preparation (per reaction):")
    output.append("")
    output.append(
        f"  • Prepare the Agent mix ({A:.3f} µL per reaction) as calculated and instructed by the Agent Prep Guide."
    )

    # Step 5: Final Mix (per reaction)
    F = R + A
    output.append("")
    output.append("Step 5 — Preparing the Final Mix (per reaction)")
    output.append(
        f"  • For ONE reaction, pipette {R:.3f} µL Reagent and {A:.3f} µL Agent into the final reaction tube "
        f"to make {F:.3f} µL final mix."
    )

    # Step 6: Master mixes for N samples using multiplier m
    R_total = R * m
    A_total = A * m
    F_total = F * m

    PS2XII_total = 10.0 * m
    AuNPs_total = 0.8 * m
    O_total = O * m
    H20_total = H20 * m

    output.append("")
    output.append("Step 6 — Preparing Master Mixes for N Multiplex Reactions")
    output.append("")
    output.append(
        f"  • You are planning to run N = {N} multiplex reactions. To account for pipetting loss, "
        f"prepare {m}× the per-reaction volume of each mix (R, A, and R-components)."
    )
    output.append("")
    output.append("  Reagent R MASTER MIX (for all reactions):")
    output.append(f"    - PS2XII: 10.0 µL × {m} = {PS2XII_total:.3f} µL")
    output.append(f"    - AuNPs: 0.8 µL × {m} = {AuNPs_total:.3f} µL")
    output.append(f"    - MULTIPLEX oligos from Probe-Primer Mix: {O:.3f} µL × {m} = {O_total:.3f} µL")
    output.append(f"    - H₂O: {H20:.3f} µL × {m} = {H20_total:.3f} µL")
    output.append(f"    → Total Reagent R master mix volume: {R_total:.3f} µL")
    output.append("")
    output.append("  Agent MASTER MIX:")
    output.append(f"    - Agent volume per reaction: {A:.3f} µL")
    output.append(f"    - Prepare: {A:.3f} µL × {m} = {A_total:.3f} µL of Agent mix in total.")
    output.append("")
    output.append("  Final Mix TOTAL VOLUME prepared:")
    output.append(f"    - Total Master Mix Prepared: {F_total:.3f} µL")
    output.append(
        f"    - You prepare this by combining the Reagent Master Mix volume ({R_total:.3f} µL) "
        f"and the Agent Master Mix volume ({A_total:.3f} µL)."
    )
    output.append("")
    output.append(
        f"  → From these master mixes, dispense {R:.3f} µL of Reagent and {A:.3f} µL of Agent into each tube "
        f"to set up N = {N} multiplex reactions."
    )
    output.append("")
    output.append("YOU ARE GOOD TO GO :) !")

    numeric = {
        "R_uL": float(R),
        "A_uL": float(A),
        "F_uL": float(F),
        "O_uL": float(O),
        "H2O_in_R_uL": float(H20),
        "N_samples": int(N),
        "multiplier": float(m),
        "L_plexes": int(L),
        "X_list": str(X_vals),
        "targets": str(target_list),
        "dyes": str(dye_list),
        "sum_J_uL": float(sum_J),
        "water_mix_base_uL": float(water_mix),
        "probe_mix_total_base_uL": float(mix_total_base),
        "probe_mix_total_scaled_uL": float(mix_total_scaled),
        "water_scaled_uL": float(water_scaled),
        "R_total_uL": float(R_total),
        "A_total_uL": float(A_total),
        "F_total_uL": float(F_total),
        "PS2XII_total_uL": float(PS2XII_total),
        "AuNPs_total_uL": float(AuNPs_total),
        "O_total_uL": float(O_total),
        "H2O_total_uL": float(H20_total),
    }

    output_text = "\n".join(output)
    return output_text, numeric


# ---------- SAVE TO EXCEL ----------
def save_to_excel(run_info, inputs, numeric_results, output_text):
    data = {
        "Timestamp": datetime.now(),
        "Batch Number": run_info.get("run_number"),
        "Date": run_info.get("date"),
        "Description": run_info.get("description"),
        "Function": inputs.get("function"),
        "Output_Text": output_text,
    }

    for k, v in inputs.items():
        if isinstance(v, (int, float)):
            data[f"input_{k}"] = v
        else:
            try:
                data[f"input_{k}"] = float(v)
            except Exception:
                data[f"input_{k}"] = v

    if isinstance(numeric_results, dict):
        for k, v in numeric_results.items():
            if isinstance(v, (int, float)) or v is None:
                data[f"num_{k}"] = v
            else:
                try:
                    data[f"num_{k}"] = float(v)
                except Exception:
                    data[f"num_{k}"] = str(v)

    df = pd.DataFrame([data])

    if not os.path.exists(excel_file):
        df.to_excel(excel_file, index=False)
    else:
        existing = pd.read_excel(excel_file)
        combined = pd.concat([existing, df], ignore_index=True, sort=False)
        combined.to_excel(excel_file, index=False)

    messagebox.showinfo("Saved", f"Run details saved to '{excel_file}'.")


# ---------- MAIN CALCULATE DISPATCHER ----------
def calculate():
    result_text.configure(state="normal")
    result_text.delete("1.0", "end")

    func = func_var.get()

    try:
        run_info = {
            "run_number": run_number_entry.get(),
            "date": run_date_entry.get(),
            "description": run_desc_entry.get("1.0", "end").strip(),
        }

        if func == "Agent Prep Guide":
            # Get desired copies
            if custom_copy_var.get():
                desired = custom_copy_entry.get()
            else:
                desired = copy_dropdown.get()

            A_agent = agent_A_entry.get()

            if not A_agent:
                raise ValueError("Please enter the Agent volume A (µL) for the Agent Prep Guide.")

            dilution_type = dilution_var.get()

            output_text, numeric_results = agent_preparation(desired, A_agent, dilution_type)

            result_text.insert("end", output_text)

            save_btn.configure(state="normal")
            inputs = {
                "function": "Agent Prep Guide",
                "desired_copies": desired,
                "A_agent_uL": A_agent,
                "dilution_type": dilution_type,
            }

        else:
            plex_choice = plex_var.get()

            if plex_choice == "Single Plex":
                R = R_entry.get()
                A = A_entry.get()
                O = O_in_R_entry.get()
                X = X_conc_entry.get()
                N_samp = N_entry.get()
                target_choice = target_var.get()
                dye_choice = dye_var.get()

                output_text, numeric_results = final_sample_preparation_new(
                    R, A, O, X, target_choice, dye_choice, N_samp
                )

                result_text.insert(
                    "end",
                    f"===== FINAL SAMPLE PREPARATION RESULTS - Single Plex ({dye_choice} {target_choice}) =====\n\n"
                )
                result_text.insert("end", output_text)

                save_btn.configure(state="normal")
                inputs = {
                    "function": "Final Sample Preparation",
                    "plex_type": plex_choice,
                    "R": R,
                    "A": A,
                    "O": O,
                    "X": X,
                    "N_samples": N_samp,
                    "target": target_choice,
                    "dye": dye_choice,
                }

            elif plex_choice == "Multiplex":
                R = R_entry.get()
                A = A_entry.get()
                O = O_in_R_entry.get()
                N_samp = N_entry.get()
                L = int(L_var.get())

                X_list = []
                t_list = []
                d_list = []
                for i in range(L):
                    X_val = multi_X_entries[i].get()
                    if str(X_val).strip() == "":
                        raise ValueError(f"X (µM) for pair {i+1} is empty.")
                    X_list.append(X_val)
                    t_list.append(multi_target_vars[i].get())
                    d_list.append(multi_dye_vars[i].get())

                output_text, numeric_results = final_sample_preparation_multiplex(
                    R, A, O, X_list, t_list, d_list, N_samp
                )

                result_text.insert(
                    "end",
                    f"===== FINAL SAMPLE PREPARATION RESULTS - Multiplex (L = {L}) =====\n\n"
                )
                result_text.insert("end", output_text)

                save_btn.configure(state="normal")
                inputs = {
                    "function": "Final Sample Preparation",
                    "plex_type": plex_choice,
                    "R": R,
                    "A": A,
                    "O": O,
                    "N_samples": N_samp,
                    "L_plexes": L,
                    "X_list": str(X_list),
                    "targets": str(t_list),
                    "dyes": str(d_list),
                }
            else:
                result_text.insert("end", "Unknown plex type.\n")
                numeric_results = {}
                inputs = {"function": "Final Sample Preparation", "plex_type": plex_choice}
                save_btn.configure(state="disabled")

        result_text.configure(state="disabled")

        def save_callback():
            save_to_excel(run_info, inputs, numeric_results, output_text)

        save_btn.configure(command=save_callback)

    except Exception as e:
        result_text.insert("end", f" ERROR: {str(e)}\n")
        result_text.configure(state="disabled")
        save_btn.configure(state="disabled")


# ------------------ UI ------------------
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Integrated Sample Prep Processor Tool")
app.geometry("1000x800")

canvas = tk.Canvas(app, bg="#eef7ff", highlightthickness=0)
scrollbar = tk.Scrollbar(app, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar.set)

scrollbar.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

main_bg = ctk.CTkFrame(canvas, fg_color="#eef7ff")
window_id = canvas.create_window((0, 0), window=main_bg, anchor="nw")


def on_frame_configure(event):
    canvas.configure(scrollregion=canvas.bbox("all"))


main_bg.bind("<Configure>", on_frame_configure)


def on_canvas_configure(event):
    canvas.itemconfig(window_id, width=event.width)


canvas.bind("<Configure>", on_canvas_configure)


def _on_mousewheel(event):
    if event.delta:
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    elif event.num in (4, 5):
        if event.num == 4:
            canvas.yview_scroll(-1, "units")
        else:
            canvas.yview_scroll(1, "units")


app.bind_all("<MouseWheel>", _on_mousewheel)
app.bind_all("<Button-4>", _on_mousewheel)
app.bind_all("<Button-5>", _on_mousewheel)

# Title
title = ctk.CTkLabel(
    main_bg,
    text="Integrated Sample Prep Processor Tool",
    font=ctk.CTkFont(size=32, weight="bold"),
    text_color="#003f63"
)
title.pack(pady=(20, 10))

# Run Information
run_frame = ctk.CTkFrame(main_bg, fg_color="#ffffff", corner_radius=12)
run_frame.pack(padx=20, pady=10, fill="x")

section_title = ctk.CTkLabel(
    run_frame,
    text="Run Information",
    font=ctk.CTkFont(size=18, weight="bold"),
    text_color="#003f63"
)
section_title.grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 5))

ctk.CTkLabel(run_frame, text="Batch Number:").grid(row=1, column=0, padx=10, pady=6, sticky="w")
run_number_entry = ctk.CTkEntry(run_frame, placeholder_text="e.g., Run-001", width=200)
run_number_entry.grid(row=1, column=1, padx=10, pady=6)

ctk.CTkLabel(run_frame, text="Date (YYYY-MM-DD):").grid(row=1, column=2, padx=10, pady=6, sticky="w")
run_date_entry = ctk.CTkEntry(run_frame, placeholder_text="e.g., 2025-11-12", width=200)
run_date_entry.insert(0, datetime.today().strftime("%Y-%m-%d"))
run_date_entry.grid(row=1, column=3, padx=10, pady=6)

ctk.CTkLabel(run_frame, text="Run Description / Notes:").grid(row=2, column=0, padx=10, pady=6, sticky="nw")
run_desc_entry = ctk.CTkTextbox(run_frame, width=600, height=80)
run_desc_entry.grid(row=2, column=1, columnspan=3, padx=10, pady=6, sticky="w")

# Function Selection
func_frame = ctk.CTkFrame(main_bg, fg_color="#ffffff", corner_radius=12)
func_frame.pack(pady=10, padx=20, fill="x")

ctk.CTkLabel(
    func_frame,
    text="Select Calculation Type",
    font=ctk.CTkFont(size=18, weight="bold"),
    text_color="#003f63"
).grid(row=0, column=0, padx=10, pady=8, sticky="w")

func_var = tk.StringVar(value="Agent Prep Guide")
func_dropdown = ctk.CTkOptionMenu(
    func_frame,
    values=["Agent Prep Guide", "Final Sample Preparation"],
    variable=func_var,
    width=250
)
func_dropdown.grid(row=0, column=1, padx=10, pady=8, sticky="w")

# Input Panel
input_frame = ctk.CTkFrame(main_bg, fg_color="#ffffff", corner_radius=12)
input_frame.pack(padx=20, pady=10, fill="x")

copy_dropdown = ctk.CTkOptionMenu(input_frame, values=copy_options_str)
custom_copy_var = tk.BooleanVar(value=False)
custom_copy_check = ctk.CTkCheckBox(input_frame, text="Custom copy number", variable=custom_copy_var)
custom_copy_entry = ctk.CTkEntry(input_frame)

dilution_var = tk.StringVar(value="Stepwise")
dilution_dropdown = ctk.CTkOptionMenu(input_frame, values=["Stepwise", "Direct"], variable=dilution_var)

# Agent Prep Guide uses this:
agent_A_entry = ctk.CTkEntry(input_frame)

# Final sample prep shared stuff
reactions_var = tk.StringVar(value="25")
reactions_dropdown = ctk.CTkOptionMenu(input_frame, values=["25", "50", "75", "100"], variable=reactions_var)

buffer_var = tk.StringVar(value="PrimeScript 2XIII")
buffer_dropdown = ctk.CTkOptionMenu(input_frame, values=["PrimeScript 2XIII", "Rover Dx Buffer"], variable=buffer_var)

plex_var = tk.StringVar(value="Single Plex")
plex_dropdown = ctk.CTkOptionMenu(input_frame, values=["Single Plex", "Multiplex"], variable=plex_var)

target_var = tk.StringVar(value="FluA")
target_dropdown = ctk.CTkOptionMenu(input_frame, values=targets, variable=target_var)
dye_var = tk.StringVar(value="FAM")
dye_dropdown = ctk.CTkOptionMenu(input_frame, values=dyes, variable=dye_var)

R_entry = ctk.CTkEntry(input_frame)
A_entry = ctk.CTkEntry(input_frame)
O_in_R_entry = ctk.CTkEntry(input_frame)
X_conc_entry = ctk.CTkEntry(input_frame)
N_entry = ctk.CTkEntry(input_frame)

# Multiplex controls
L_var = tk.StringVar(value="2")
multi_target_vars = [tk.StringVar(value=targets[0]) for _ in range(7)]
multi_dye_vars = [tk.StringVar(value=dyes[0]) for _ in range(7)]
multi_target_dropdowns = [
    ctk.CTkOptionMenu(input_frame, values=targets, variable=multi_target_vars[i]) for i in range(7)
]
multi_dye_dropdowns = [
    ctk.CTkOptionMenu(input_frame, values=dyes, variable=multi_dye_vars[i]) for i in range(7)
]
multi_X_entries = [ctk.CTkEntry(input_frame) for _ in range(7)]


def show_inputs(*args):
    for widget in input_frame.winfo_children():
        widget.grid_forget()

    ctk.CTkLabel(
        input_frame,
        text="Input Parameters",
        font=ctk.CTkFont(size=18, weight="bold"),
        text_color="#003f63"
    ).grid(row=0, column=0, columnspan=6, padx=10, pady=(10, 5), sticky="w")

    if func_var.get() == "Agent Prep Guide":
        ctk.CTkLabel(input_frame, text="Select standard copy number (copies per reaction):").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        copy_dropdown.grid(row=1, column=1, padx=10, pady=5)
        custom_copy_check.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        custom_copy_entry.grid(row=2, column=1, padx=10, pady=5)

        ctk.CTkLabel(input_frame, text="Agent volume A (µL) per reaction:").grid(
            row=3, column=0, padx=10, pady=5, sticky="w"
        )
        agent_A_entry.grid(row=3, column=1, padx=10, pady=5, sticky="w")
        if agent_A_entry.get() == "":
            agent_A_entry.insert(0, "5")

        ctk.CTkLabel(input_frame, text="Select dilution type:").grid(
            row=4, column=0, padx=10, pady=5, sticky="w"
        )
        dilution_dropdown.grid(row=4, column=1, padx=10, pady=5)

    elif func_var.get() == "Final Sample Preparation":
        ctk.CTkLabel(input_frame, text="Select Plex Type:").grid(
            row=1, column=0, padx=10, pady=5, sticky="w"
        )
        plex_dropdown.grid(row=1, column=1, padx=10, pady=5, sticky="w")

        if plex_var.get() == "Single Plex":
            ctk.CTkLabel(input_frame, text="Number of Samples (N):").grid(
                row=2, column=0, padx=10, pady=5, sticky="w"
            )
            N_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")
            if N_entry.get() == "":
                N_entry.insert(0, "1")

            ctk.CTkLabel(input_frame, text="Select Target:").grid(
                row=3, column=0, padx=10, pady=5, sticky="w"
            )
            target_dropdown.grid(row=3, column=1, padx=10, pady=5)

            ctk.CTkLabel(input_frame, text="Select Dye:").grid(
                row=4, column=0, padx=10, pady=5, sticky="w"
            )
            dye_dropdown.grid(row=4, column=1, padx=10, pady=5)

            ctk.CTkLabel(input_frame, text="Reagent Volume R (µL) per reaction:").grid(
                row=2, column=2, padx=10, pady=5, sticky="w"
            )
            R_entry.grid(row=2, column=3, padx=100, pady=5)

            ctk.CTkLabel(input_frame, text="Agent Volume A (µL) per reaction:").grid(
                row=3, column=2, padx=10, pady=5, sticky="w"
            )
            A_entry.grid(row=3, column=3, padx=100, pady=5)

            ctk.CTkLabel(input_frame, text="Oligo in R (O) (µL) per reaction:").grid(
                row=4, column=2, padx=10, pady=5, sticky="w"
            )
            O_in_R_entry.grid(row=4, column=3, padx=100, pady=5)

            ctk.CTkLabel(input_frame, text="Desired Oligo Conc X (µM):").grid(
                row=5, column=2, padx=10, pady=5, sticky="w"
            )
            X_conc_entry.grid(row=5, column=3, padx=100, pady=5)

            ctk.CTkLabel(
                input_frame,
                text="(Final Mix F = R + A per reaction)",
                text_color="#003f63"
            ).grid(row=6, column=0, columnspan=4, padx=10, pady=(5, 10), sticky="center")

        else:  # Multiplex
            ctk.CTkLabel(input_frame, text="Number of Samples (N):").grid(
                row=2, column=0, padx=10, pady=5, sticky="w"
            )
            N_entry.grid(row=2, column=1, padx=10, pady=5, sticky="w")
            if N_entry.get() == "":
                N_entry.insert(0, "1")

            ctk.CTkLabel(input_frame, text="Number of Plexes (L):").grid(
                row=2, column=2, padx=10, pady=5, sticky="w"
            )
            L_option = ctk.CTkOptionMenu(input_frame, values=[str(i) for i in range(1, 8)], variable=L_var)
            L_option.grid(row=2, column=3, padx=10, pady=5, sticky="w")

            L_num = int(L_var.get())
            start_row = 3
            ctk.CTkLabel(input_frame, text="Pair").grid(row=start_row, column=0, padx=10, pady=5, sticky="w")
            ctk.CTkLabel(input_frame, text="Target").grid(row=start_row, column=1, padx=10, pady=5, sticky="w")
            ctk.CTkLabel(input_frame, text="Dye").grid(row=start_row, column=2, padx=10, pady=5, sticky="w")
            ctk.CTkLabel(input_frame, text="X (µM)").grid(row=start_row, column=3, padx=10, pady=5, sticky="w")

            for i in range(L_num):
                row = start_row + 1 + i
                ctk.CTkLabel(input_frame, text=f"{i+1}").grid(row=row, column=0, padx=10, pady=5, sticky="w")
                multi_target_dropdowns[i].grid(row=row, column=1, padx=10, pady=5, sticky="w")
                multi_dye_dropdowns[i].grid(row=row, column=2, padx=10, pady=5, sticky="w")
                multi_X_entries[i].grid(row=row, column=3, padx=10, pady=5, sticky="w")

            base_row = start_row + 1 + L_num + 1
            ctk.CTkLabel(input_frame, text="Reagent Volume R (µL) per reaction:").grid(
                row=base_row, column=0, padx=10, pady=5, sticky="w"
            )
            R_entry.grid(row=base_row, column=1, padx=10, pady=5, sticky="w")

            ctk.CTkLabel(input_frame, text="Agent Volume A (µL) per reaction:").grid(
                row=base_row+1, column=0, padx=10, pady=5, sticky="w"
            )
            A_entry.grid(row=base_row+1, column=1, padx=10, pady=5, sticky="w")

            ctk.CTkLabel(input_frame, text="Oligo in R (O) (µL) per reaction:").grid(
                row=base_row+2, column=0, padx=10, pady=5, sticky="w"
            )
            O_in_R_entry.grid(row=base_row+2, column=1, padx=10, pady=5, sticky="w")

            ctk.CTkLabel(
                input_frame,
                text="(Final Mix F = R + A per reaction)",
                text_color="#003f63"
            ).grid(row=base_row+3, column=0, columnspan=4, padx=10, pady=(5, 10), sticky="center")


func_var.trace("w", show_inputs)
plex_var.trace("w", show_inputs)
L_var.trace("w", show_inputs)
show_inputs()

# Buttons
button_frame = ctk.CTkFrame(main_bg, fg_color="#eef7ff")
button_frame.pack(pady=(5, 0))

calc_btn = ctk.CTkButton(
    button_frame,
    text="Calculate",
    command=calculate,
    width=200,
    height=45,
    font=ctk.CTkFont(size=18, weight="bold")
)
calc_btn.grid(row=0, column=0, padx=20, pady=15)

save_btn = ctk.CTkButton(
    button_frame,
    text="Save Run",
    state="disabled",
    width=200,
    height=45,
    font=ctk.CTkFont(size=18, weight="bold")
)
save_btn.grid(row=0, column=1, padx=20, pady=15)

# Output Panel
output_frame = ctk.CTkFrame(main_bg, fg_color="#ffffff", corner_radius=30)
output_frame.pack(fill="both", expand=True, padx=20, pady=20)

ctk.CTkLabel(
    output_frame,
    text="Output",
    font=ctk.CTkFont(size=18, weight="bold"),
    text_color="#003f63"
).pack(anchor="w", padx=10, pady=(10, 5))

result_text = scrolledtext.ScrolledText(
    output_frame,
    wrap="word",
    font=("Arial", 14),
    bg="#e3f4ff",
    fg="#000",
    insertbackground="black",
    relief="flat",
    height=20
)
result_text.pack(fill="both", expand=True, padx=10, pady=10)
result_text.configure(state="disabled")

app.update_idletasks()
canvas.configure(scrollregion=canvas.bbox("all"))

app.mainloop()
