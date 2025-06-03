import pandas as pd
import matplotlib.pyplot as plt

# 📥 Load Excel data
file_path = "input_data.xlsx"
df_gu13 = pd.read_excel(file_path, sheet_name="GU13")
df_tm4 = pd.read_excel(file_path, sheet_name="TM4")
df_gu13.columns = df_gu13.columns.str.strip()
df_tm4.columns = df_tm4.columns.str.strip()

# 🔁 GWP calculation function per component
def calculate_component_gwp(df, product_name):
    limestone = cement = clay = rc_mix = 0
    component_gwps = {}

    for _, row in df.iterrows():
        name = str(row["Name"]).strip().lower().replace("–", "-").replace("—", "-")
        try:
            amount = float(row["Amount"])
        except:
            continue

        if name == "limestone":
            limestone = amount
        elif name == "cem i 52.5 r":
            cement = amount
        elif name == "concresol / clay powder":
            clay = amount
        elif "rc" in name and "mix" in name:
            rc_mix = amount

    binder_mass_ton = (cement + clay + limestone) / 1000
    rc_mix_mass_ton = rc_mix / 1000

    for _, row in df.iterrows():
        if pd.isna(row["Amount"]) or pd.isna(row["CO2 Factor"]):
            continue

        name = str(row["Name"]).strip()
        name_lower = name.lower()
        category = str(row["Category"]).strip().lower()
        try:
            amount = float(row["Amount"])
            co2 = float(row["CO2 Factor"])
        except:
            continue

        if name_lower == "mixing plant":
            continue

        # GWP calculation
        if category == "transport":
            if name_lower == "cement transport":
                gwp = amount * binder_mass_ton * co2
                name = "Binder Transport"
            elif "recycled" in name_lower:
                gwp = amount * rc_mix_mass_ton * co2
                name = "Recycled Aggregate Transport"
            else:
                gwp = 0
        else:
            gwp = amount * co2

        component_gwps[name] = gwp

    return pd.Series(component_gwps, name=product_name)

# ✅ Calculate component-wise GWP
gu13_series = calculate_component_gwp(df_gu13, "GU13 Concrete")
tm4_series = calculate_component_gwp(df_tm4, "TM4 Excavated Soil")

# 📊 Component-wise comparison chart
comparison_df = pd.concat([gu13_series, tm4_series], axis=1).fillna(0)
comparison_df = comparison_df.sort_index()

ax = comparison_df.plot(kind="bar", figsize=(12, 6), color=["#1f77b4", "#ff7f0e"])
plt.title("Component-wise GWP Comparison")
plt.ylabel("GWP (kg CO₂-eq)")
plt.xlabel("Material / Process")
plt.xticks(rotation=45, ha="right")
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.legend(title="Product")
plt.tight_layout()
plt.show()

# ✅ Total GWP values
total_gu13 = gu13_series.sum()
total_tm4 = tm4_series.sum()

# 📊 Total GWP comparison chart
products = ["GU13", "TM4"]
values = [total_gu13, total_tm4]
colors = ["#1f77b4", "#ff7f0e"]

plt.figure(figsize=(8, 5))
bars = plt.bar(products, values, color=colors)

for i, bar in enumerate(bars):
    yval = values[i]
    plt.text(bar.get_x() + bar.get_width() / 2, yval * 1.02, f"{yval:.2f} kg CO₂-eq",
             ha='center', va='bottom', fontsize=10, fontweight='bold')

plt.title("Total GWP Comparison")
plt.ylabel("GWP (kg CO₂-eq)")
plt.ylim(0, max(values) * 1.3)
plt.grid(axis="y", linestyle="--", alpha=0.6)
plt.tight_layout()
plt.show()
