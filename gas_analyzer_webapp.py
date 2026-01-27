=== GAS TURBINE FUEL ANALYSIS TOOL ===

INSTRUCTIONS:
1. Enter mol% for each component in Column B (User Input) OR
2. Use Pipeline Spec values shown below for quick calculations
3. Calculated properties appear in the RESULTS section below
4. Total mol% must equal 100%
5. Results shown in both SI and US units

---COMPONENT DATABASE---
Component | Mol% Input | Mol Weight | LHV (MJ/kg) | HHV (MJ/kg) | Sp. Gravity
          |     (B)    |    (C)     |     (D)     |     (E)     |     (F)
----------|------------|------------|-------------|-------------|-------------
Methane (CH4) | | 16.043 | 50.01 | 55.50 | 0.5539
Ethane (C2H6) | | 30.070 | 47.49 | 51.88 | 1.0382
Propane (C3H8) | | 44.097 | 46.35 | 50.36 | 1.5225
n-Butane (C4H10) | | 58.123 | 45.75 | 49.50 | 2.0068
i-Butane (C4H10) | | 58.123 | 45.61 | 49.36 | 2.0068
n-Pentane (C5H12) | | 72.150 | 45.36 | 49.01 | 2.4911
i-Pentane (C5H12) | | 72.150 | 45.24 | 48.89 | 2.4911
n-Hexane (C6H14) | | 86.177 | 45.10 | 48.68 | 2.9754
Heptane (C7H16) | | 100.204 | 44.93 | 48.45 | 3.4597
Octane (C8H18) | | 114.231 | 44.79 | 48.26 | 3.9440
Nonane (C9H20) | | 128.258 | 44.68 | 48.10 | 4.4283
Decane (C10H22) | | 142.285 | 44.60 | 47.98 | 4.9126
Hydrogen (H2) | | 2.016 | 120.00 | 141.80 | 0.0696
Carbon Monoxide (CO) | | 28.010 | 10.10 | 10.10 | 0.9671
Carbon Dioxide (CO2) | | 44.010 | 0.00 | 0.00 | 1.5196
Nitrogen (N2) | | 28.014 | 0.00 | 0.00 | 0.9672
Oxygen (O2) | | 31.999 | 0.00 | 0.00 | 1.1048
Hydrogen Sulfide (H2S) | | 34.081 | 15.20 | 16.53 | 1.1766
Water (H2O) | | 18.015 | 0.00 | 0.00 | 0.6219
Helium (He) | | 4.003 | 0.00 | 0.00 | 0.1382
Argon (Ar) | | 39.948 | 0.00 | 0.00 | 1.3788

PIPELINE SPEC (Typical Natural Gas):
Methane: 95.0% | Ethane: 2.5% | Propane: 0.5%
n-Butane: 0.2% | CO2: 1.0% | N2: 0.8%

---RESULTS SECTION (Starting Row 25)---

EXCEL FORMULAS FOR COLUMN C (SI Units):

Row 25 - Total Composition Check:
Cell C25: =SUM(B2:B22)
Label: "Total mol%"
Target: Must = 100.0%

Row 26 - Average Molecular Weight:
Cell C26: =SUMPRODUCT(B2:B22/100,C2:C22)
Label: "Molecular Weight"
Unit: g/mol (or lb/lbmol)

Row 27 - Specific Gravity:
Cell C27: =C26/28.97
Label: "Specific Gravity"
Unit: dimensionless (same for US)

Row 28 - Gas Density at STP:
Cell C28: =C26/22.414
Label: "Gas Density (STP)"
Unit (SI): kg/m³

Row 29 - LHV Mass Basis:
Cell C29: =SUMPRODUCT((B2:B22/100*C2:C22)/C$26,D2:D22)
Label: "LHV (mass basis)"
Unit (SI): MJ/kg

Row 30 - LHV Volume Basis:
Cell C30: =C29*(C26/22.414)*1000
Label: "LHV (volume basis)"
Unit (SI): MJ/m³

Row 31 - HHV Mass Basis:
Cell C31: =SUMPRODUCT((B2:B22/100*C2:C22)/C$26,E2:E22)
Label: "HHV (mass basis)"
Unit (SI): MJ/kg

Row 32 - HHV Volume Basis:
Cell C32: =C31*(C26/22.414)*1000
Label: "HHV (volume basis)"
Unit (SI): MJ/m³

Row 33 - Wobbe Index Lower:
Cell C33: =C30/SQRT(C27)
Label: "Wobbe Index (Lower)"
Unit (SI): MJ/m³

Row 34 - Wobbe Index Higher:
Cell C34: =C32/SQRT(C27)
Label: "Wobbe Index (Higher)"
Unit (SI): MJ/m³

Row 35 - Hydrogen Content:
Cell C35: =B14
Label: "H2 Content"
Unit: mol% (assuming H2 is in row 14)

Row 36 - Inert Content:
Cell C36: =B16+B17
Label: "CO2 + N2 Content"
Unit: mol% (assuming CO2 in row 16, N2 in row 17)

Row 37 - H2S Content:
Cell C37: =B19*10000
Label: "H2S Content"
Unit: ppmv (assuming H2S is in row 19)

---US UNIT CONVERSIONS (Column D)---

EXCEL FORMULAS FOR COLUMN D (US Units):

Row 26 - Molecular Weight:
Cell D26: =C26
Unit: lb/lbmol (same value)

Row 27 - Specific Gravity:
Cell D27: =C27
Unit: dimensionless (same value)

Row 28 - Gas Density:
Cell D28: =C28*0.062428
Unit: lb/ft³

Row 29 - LHV Mass Basis:
Cell D29: =C29*429.923
Unit: Btu/lb

Row 30 - LHV Volume Basis:
Cell D30: =C30*26.839
Unit: Btu/scf

Row 31 - HHV Mass Basis:
Cell D31: =C31*429.923
Unit: Btu/lb

Row 32 - HHV Volume Basis:
Cell D32: =C32*26.839
Unit: Btu/scf

Row 33 - Wobbe Index Lower:
Cell D33: =C33*26.839
Unit: Btu/scf

Row 34 - Wobbe Index Higher:
Cell D34: =C34*26.839
Unit: Btu/scf

Row 35-37 - Same values:
Cells D35, D36, D37: =C35, =C36, =C37
(mol% and ppmv are same in both systems)

---CONVERSION FACTORS REFERENCE---
1 MJ/kg = 429.923 Btu/lb
1 MJ/m³ = 26.839 Btu/ft³ (or Btu/scf)
1 kg/m³ = 0.062428 lb/ft³
scf = standard cubic feet (60°F, 14.7 psia)
STP = 0°C, 1 atm

---PIPELINE SPEC CALCULATED RESULTS---

Property | SI Units | US Units
---------|----------|----------
Total Composition | 100.0 | mol%
Molecular Weight | 16.993 | 16.993 g/mol | lb/lbmol
Specific Gravity | 0.587 | 0.587
Gas Density (STP) | 0.758 kg/m³ | 0.0473 lb/ft³
LHV (mass basis) | 48.95 MJ/kg | 21,049 Btu/lb
LHV (volume basis) | 37.12 MJ/m³ | 996 Btu/scf
HHV (mass basis) | 54.24 MJ/kg | 23,322 Btu/lb
HHV (volume basis) | 41.13 MJ/m³ | 1,104 Btu/scf
Wobbe Index (Lower) | 48.46 MJ/m³ | 1,301 Btu/scf
Wobbe Index (Higher) | 53.70 MJ/m³ | 1,441 Btu/scf
H2 Content | 0.0 mol% | 0.0 mol%
CO2 + N2 Content | 1.8 mol% | 1.8 mol%
H2S Content | 0 ppmv | 0 ppmv

---TURBINE SUITABILITY ASSESSMENT---

TYPICAL GAS TURBINE SPECIFICATIONS:

Parameter | SI Limit | US Limit | Pipeline Result (SI / US) | Status
----------|----------|----------|---------------------------|--------
Wobbe Index (Lower) | 47-51 MJ/m³ | 1,260-1,370 Btu/scf | 48.46 MJ/m³ / 1,301 Btu/scf | ✓ Good
Wobbe Index (Higher) | 52-57 MJ/m³ | 1,400-1,520 Btu/scf | 53.70 MJ/m³ / 1,441 Btu/scf | ✓ Good
LHV (volume basis) | 32-40 MJ/m³ | 850-1,050 Btu/scf | 37.12 MJ/m³ / 996 Btu/scf | ✓ Good
HHV (volume basis) | 35-45 MJ/m³ | 950-1,150 Btu/scf | 41.13 MJ/m³ / 1,104 Btu/scf | ✓ Good
LHV (mass basis) | 45-52 MJ/kg | 19,000-22,000 Btu/lb | 48.95 MJ/kg / 21,049 Btu/lb | ✓ Good
HHV (mass basis) | 50-58 MJ/kg | 21,000-25,000 Btu/lb | 54.24 MJ/kg / 23,322 Btu/lb | ✓ Good
Specific Gravity | 0.55-0.75 | 0.55-0.75 | 0.587 / 0.587 | ✓ Good
Gas Density (STP) | 0.65-0.90 kg/m³ | 0.04-0.06 lb/ft³ | 0.758 kg/m³ / 0.047 lb/ft³ | ✓ Good
H2 Content | <5 mol% | <5 mol% | 0.0 mol% | ✓ Excellent
CO2 + N2 (Inerts) | <20 mol% | <20 mol% | 1.8 mol% | ✓ Excellent
H2S Content | <5 ppmv | <5 ppmv (<¼ grain/100scf) | 0 ppmv | ✓ Excellent
Total Sulfur | <20 ppmv | <20 ppmv | - | Check if needed
Molecular Weight | 16-20 g/mol | 16-20 lb/lbmol | 16.993 | ✓ Good

SUITABILITY CHECKLIST:
□ Total mol% = 100%
□ Wobbe Index within ±5% of turbine design value
□ H2 content acceptable (flashback risk if >5%)
□ Inert content not excessive (affects stability)
□ H2S within limits (corrosion protection)
□ Heating value adequate for power requirements

NOTES:
- Always consult turbine OEM specifications for exact limits
- Consider fuel gas conditioning if out of spec
- Monitor combustion dynamics for unusual compositions
- Review derating requirements for off-spec fuels
