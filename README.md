# MSc Downward Counterfactual Analysis

Downward counterfactual analysis of the 2023 Türkiye-Syria earthquakes, comparing two fatality-estimation frameworks — **HAZUS** and **GEM** — across two "what if" questions:

- **SQ1 (timing):** how would the death toll have differed if the earthquake had struck at a different time of day (04:17 night-time vs. 14:00 day-time vs. 17:00 transit hours)?
- **SQ2 (compliance):** how would the death toll have differed under a uniform building-code compliance level (Pre-Code / Low-Code / High-Code), instead of the actual 2023 mixed-compliance building stock?

The two questions are also explored jointly via a **retrofit sensitivity sweep**, simulating a gradual transition from the real 2023 compliance mix (0% retrofit) to full High-Code compliance (100% retrofit), for both models and all three timing scenarios.

## Data requirements

The script auto-detects its input workbook: it scans every `.xlsx` file in the same folder and picks the first one containing a sheet named `Fragility PGA`, so the workbook doesn't need an exact filename. That workbook must contain the following sheets:

- `Fragility PGA`
- `Collapse Rate`
- `Casualty Severity 4`
- `Summary Model Input`
- `Fatalities Data (Long)` (GEM fatality lookup tables)
- `RAW Floor Distribution`

## Requirements

```
openpyxl
numpy
scipy
pandas
matplotlib
geopandas
shapely
requests
```

Internet access is required for two one-time downloads (Turkey province boundaries and GEM's global active-faults dataset), used to build the choropleth maps.

## Running

```bash
python3 MSc_Downward_Counterfactual_Analysis.py
```

Outputs are written to the script's own directory.

## Outputs

**Charts (PNG)**
- 3x3 compliance x timing matrices — national totals and per-province heatmaps, HAZUS and GEM
- SQ1 vs. SQ2 sensitivity ratio charts, and their interaction/tornado breakdowns
- Retrofit sensitivity sweeps (0-100% retrofit coverage), national and per-province, all three timings
- Fatality-reduction-vs-retrofit comparison, HAZUS vs. GEM
- HAZUS vs. GEM comparison charts (national totals, compliance benefit captured, factual deaths by province/timing)
- Study area map and province-level compliance/retrofit-ceiling maps

**Tables (XLSX)**
- `HAZUS_3x3_Matrix_Results.xlsx` — national totals + per-province detail
- `retrofit_sweep_results.xlsx`, `retrofit_sweep_hazus_all_times.xlsx`
- `retrofit_sweep_gem_all_times.xlsx`, `retrofit_sweep_gem_by_province.xlsx`

## Method summary

Fatality estimates follow the standard HAZUS damage-chain: for each building type and PGA level, a log-normal fragility curve gives the probability of complete structural damage, multiplied by a collapse rate and a Severity-4 (fatality) casualty rate. The GEM estimates follow the same occupancy/exposure structure but use GEM's empirical loss-ratio curves in place of HAZUS fragility functions. Both models are run across 11 earthquake-affected provinces, three building-code eras (Pre-Code / Low-Code / High-Code), and three event timings, and validated against the actual reported death toll (Chen et al., 2025 / MoEUCC: 52,724).
