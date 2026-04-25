"""
create_sample_excel.py
======================
Run this script ONCE to generate a sample 'medicines.xlsx' file
that you can fill in with your actual medicine data.

Usage:
    python create_sample_excel.py
"""

import pandas as pd

sample_data = {
    "Name": [
        "Paracetamol 500mg",
        "Amoxicillin 250mg",
        "Ibuprofen 400mg",
        "Metformin 500mg",
        "Atorvastatin 10mg",
    ],
    "Generic Name": [
        "Paracetamol",
        "Amoxicillin",
        "Ibuprofen",
        "Metformin",
        "Atorvastatin",
    ],
    "Brand Name": [
        "Panadol",
        "Amoxil",
        "Brufen",
        "Glucophage",
        "Lipitor",
    ],
    "Manufacturer": [
        "GSK",
        "Pfizer",
        "Abbott",
        "Merck",
        "Pfizer",
    ],
    "Category": [
        "Analgesic",
        "Antibiotic",
        "NSAID",
        "Antidiabetic",
        "Statin",
    ],
    "Dosage Form": [
        "Tablet",
        "Capsule",
        "Tablet",
        "Tablet",
        "Tablet",
    ],
    "Strength": [
        "500 mg",
        "250 mg",
        "400 mg",
        "500 mg",
        "10 mg",
    ],
    "Unit Price": [0.05, 0.25, 0.10, 0.15, 0.50],
    "Stock Quantity": [1000, 500, 750, 600, 300],
    "Expiry Date": [
        "2027-06-30",
        "2026-12-31",
        "2027-03-31",
        "2028-01-31",
        "2027-09-30",
    ],
    "Description": [
        "Used for pain relief and fever reduction.",
        "Broad-spectrum antibiotic for bacterial infections.",
        "Anti-inflammatory, analgesic, and antipyretic.",
        "First-line treatment for type 2 diabetes.",
        "Lowers cholesterol and triglycerides.",
    ],
    "Side Effects": [
        "Nausea, rash (rare)",
        "Diarrhea, rash, allergic reaction",
        "Stomach upset, heartburn",
        "Nausea, diarrhea, vitamin B12 deficiency",
        "Muscle pain, liver enzyme elevation",
    ],
    "Contraindications": [
        "Severe liver disease",
        "Penicillin allergy",
        "Peptic ulcer, renal impairment",
        "Renal impairment, liver disease",
        "Liver disease, pregnancy",
    ],
    "Storage Conditions": [
        "Store below 25°C, dry place",
        "Store below 25°C, dry place",
        "Store below 30°C",
        "Store below 25°C",
        "Store below 25°C",
    ],
}

df = pd.DataFrame(sample_data)
output_file = "medicines.xlsx"
df.to_excel(output_file, index=False, sheet_name="Sheet1")
print(f"✅  Sample Excel file created: {output_file}")
print(f"   Rows: {len(df)}")
print("\nColumn headers used:")
for col in df.columns:
    print(f"   - {col}")
print(
    "\nEdit this file with your real medicine data, "
    "then run:\n    python medicine_autobot.py"
)
