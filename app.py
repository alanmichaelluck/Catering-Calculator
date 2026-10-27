import streamlit as st
import openpyxl
import pandas as pd
import re
from datetime import datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="Catering + Prep Report Integrator", layout="wide")
st.title("📋 Unified Prep Report & Multi-Day Catering Calculator")

# ==============================================================================
# 1. CROSSWALK MAPPING: CATERING CALCULATOR ROW -> PREP REPORT RECIPE(S)
# ==============================================================================
CATERING_CROSSWALK = {
    # --- SIDES ---
    4: [  # Mac
        {"prep_recipe": "Cheese Sauce (Flower)", "source_col": "F", "multiplier": 1.0},
        {"prep_recipe": "Gluten Free Pasta (Flower)", "source_col": "F", "multiplier": 0.25},
        {"prep_recipe": "Gouda Cheese Blend (Flower)", "source_col": "F", "multiplier": 0.125},
        {"prep_recipe": "Bread Crumbs (Flower)", "source_col": "F", "multiplier": 0.0625},
    ],
    5: [  # Quinoa Corn
        {"prep_recipe": "Creamed Corn Prep (Flower)", "source_col": "F", "multiplier": 1.0},
        {"prep_recipe": "Grains - Red Quinoa Prepped (Flower)", "source_col": "E", "multiplier": 0.25},
    ],
    6: [{"prep_recipe": "Roasted Sweet Potatoes (Flower)", "source_col": "F", "multiplier": 1.0}], # Sweets
    7: [{"prep_recipe": "Smashed Potatoes Prep (Flower)", "source_col": "F", "multiplier": 2.0}],  # Smash
    8: [  # Broccoli
        {"prep_recipe": "Roasted Broccoli Florets - Prep (Flower)", "source_col": "E", "multiplier": 4.0},
        {"prep_recipe": "Broccoli Dressing (Flower)", "source_col": "F", "multiplier": 0.25},
    ],
    9: [{"prep_recipe": "Fruit Salad Mix (Flower)", "source_col": "E", "multiplier": 2.0}],      # Fruit
    10: [{"prep_recipe": "Brown Rice Prep (Flower)", "source_col": "E", "multiplier": 0.5}],       # Brown Rice
    11: [{"prep_recipe": "Jasmine Rice Prep (Flower)", "source_col": "E", "multiplier": 2.0}],     # Jasmine Rice
    12: [{"prep_recipe": "Butternut & Yukon Gratin (Flower)", "source_col": "F", "multiplier": 5.0}], # Butternut Gratin
    13: [{"prep_recipe": "Salad Mix - Brussels & Kale (Flower)", "source_col": "E", "multiplier": 2.0}], # Kale Salad
    14: [{"prep_recipe": "Cauliflower Risotto (Flower)", "source_col": "F", "multiplier": 5.0}],    # Cauli Risotto
    16: [{"prep_recipe": "Roasted Brussels Sprouts (Flower)", "source_col": "E", "multiplier": 2.0}], # Brussels Sprouts

    # --- BOWLS ---
    18: [{"prep_recipe": "Flo Cho Chow Sauce (Flower)", "source_col": "E", "multiplier": 0.0625}], # Flo Cho Chow
    19: [  # Forbidden
        {"prep_recipe": "Black Rice Prep (Flower)", "source_col": "E", "multiplier": 0.125},
        {"prep_recipe": "Sesame Vinaigrette (Flower)", "source_col": "E", "multiplier": 0.0625},
    ],
    20: [  # Chicken Enchiladas
        {"prep_recipe": "Enchilada Sauce (Flower)", "source_col": "E", "multiplier": 0.09375},
        {"prep_recipe": "Red Chili Braised Chicken (Flower)", "source_col": "E", "multiplier": 0.125},
        {"prep_recipe": "Green Chili Queso Dip (Flower)", "source_col": "E", "multiplier": 0.05},
    ],
    21: [{"prep_recipe": "Miso Glaze (Flower)", "source_col": "E", "multiplier": 0.046875}],       # Miso Glazed Salmon
    22: [  # Chicken Yakisoba
        {"prep_recipe": "Yaki Noodle Veg Set - Flower", "source_col": "E", "multiplier": 1.0},
        {"prep_recipe": "Garlic Noodle Sauce (Flower)", "source_col": "E", "multiplier": 0.0625},
        {"prep_recipe": "Ramen Noodles - Blanched (Flower)", "source_col": "E", "multiplier": 1.0},
    ],
    23: [  # Mother Earth
        {"prep_recipe": "Supergrain Mix (Flower)", "source_col": "F", "multiplier": 0.5},
        {"prep_recipe": "Red Pepper Miso Marmalade (Flower)", "source_col": "E", "multiplier": 0.046875},
        {"prep_recipe": "Broccoli Dressing (Flower)", "source_col": "E", "multiplier": 0.046875},
    ],

    # --- SALADS & DRESSINGS ---
    25: [{"prep_recipe": "Salad Mix - Skinny Cobb (Flower)", "source_col": "E", "multiplier": 1.0}],
    26: [{"prep_recipe": "Smoked Blue Cheese Vinaigrette (Flower)", "source_col": "E", "multiplier": 1.0 / 32.0}],
    27: [{"prep_recipe": "Salad Mix - Ginger Miso Crunch (Flower)", "source_col": "E", "multiplier": 1.0}],
    28: [{"prep_recipe": "Ginger Miso Vinaigrette (Flower)", "source_col": "E", "multiplier": 1.0 / 32.0}],
    29: [{"prep_recipe": "Salad Mix - Chopped Veg (Flower)", "source_col": "E", "multiplier": 1.0}],
    30: [{"prep_recipe": "Oregano Dressing (Flower)", "source_col": "E", "multiplier": 1.0 / 32.0}],
    31: [{"prep_recipe": "Avocado Chicken Caesar Salad (Flower)", "source_col": "E", "multiplier": 0.25}],
    32: [{"prep_recipe": "Avocado Caesar Dressing (Flower)", "source_col": "E", "multiplier": 1.0 / 32.0}],
    34: [{"prep_recipe": "Maple Sherry Vinaigrette (Flower)", "source_col": "E", "multiplier": 1.0 / 32.0}],

    # --- STARTERS ---
    36: [{"prep_recipe": "Avocado Hummus (Flower)", "source_col": "E", "multiplier": 0.5}],
    39: [{"prep_recipe": "Classic Hummus Prep (Flower)", "source_col": "E", "multiplier": 0.5}],

    # --- WRAPS ---
    43: [{"prep_recipe": "Turkey (Wraps) - Portion", "source_col": "E", "multiplier": 1.0}],
    44: [{"prep_recipe": "Marinated Chicken Breast Prep (Flower)", "source_col": "E", "multiplier": 0.25}],
    46: [{"prep_recipe": "Rebel Sets (Flower)", "source_col": "E", "multiplier": 1.0}],

    # --- PROTEINS ---
    48: [{"prep_recipe": "Marinated Chicken Breast Prep (Flower)", "source_col": "E", "multiplier": 1.0}],
    49: [{"prep_recipe": "Marinated Flank Steak Prep (Flower)", "source_col": "E", "multiplier": 1.0}],
    50: [{"prep_recipe": "Miso Marinated Salmon (Flower)", "source_col": "F", "multiplier": 1.0}],
    51: [{"prep_recipe": "Marinated Tofu (Flower)", "source_col": "E", "multiplier": 4.0}],
    52: [{"prep_recipe": "Marinated Shrimp - Portion (Flower)", "source_col": "E", "multiplier": 4.0}],

    # --- DESSERTS ---
    54: [
        {"prep_recipe": "Chocolate Chip Cookies - FLOWER", "source_col": "E", "multiplier": 1.0},
        {"prep_recipe": "Dessert - Chocolate Chip Cookies Dough (Flower)", "source_col": "E", "multiplier": 1.0},
    ],
    55: [
        {"prep_recipe": "Double Chocolate Almond Brownie - FLOWER", "source_col": "E", "multiplier": 1.0},
        {"prep_recipe": "Dessert - Chocolate Almond Brownie Batter (Flower)", "source_col": "E", "multiplier": 1.0},
    ],
    56: [
        {"prep_recipe": "Chocolate Chip Cookies - FLOWER", "source_col": "E", "multiplier": 0.5},
        {"prep_recipe": "Dessert - Chocolate Chip Cookies Dough (Flower)", "source_col": "E", "multiplier": 0.5},
        {"prep_recipe": "Double Chocolate Almond Brownie - FLOWER", "source_col": "E", "multiplier": 0.5},
        {"prep_recipe": "Dessert - Chocolate Almond Brownie Batter (Flower)", "source_col": "E", "multiplier": 0.5},
    ],
}

# ==============================================================================
# 2. HELPER FUNCTIONS
# ==============================================================================
def parse_shelf_life_days(sl_str):
    s = str(sl_str).strip().lower()
    if s.endswith('h'):
        return 1
    if s.endswith('d'):
        days = int(re.findall(r'\d+', s)[0])
        return min(max(days, 1), 3)
    return 1

def parse_batch_size(bs_str, iu_str):
    bs = str(bs_str).strip()
    iu = str(iu_str).strip()
    if bs == '1 X = 4" Hotel Pan':
        return 1.0
    if bs == '1 X = 0.75 Qt' and iu == 'Cup':
        return 3.0
    if '1 X =' in bs:
        after = bs.split('1 X =')[1].strip()
        m = re.match(r'^(\d+(?:\.\d+)?)', after)
        if m:
            return float(m.group(1))
    return 1.0

def extract_catering_demands(cat_file_bytes):
    wb = openpyxl.load_workbook(BytesIO(cat_file_bytes), data_only=True)
    ws = wb['Catering Calculator']
    raw_date = ws.cell(1, 8).value
    if isinstance(raw_date, datetime):
        cat_date = raw_date.date()
    else:
        cat_date = pd.to_datetime(str(raw_date)).date()

    demands = {}
    for row_idx, mappings in CATERING_CROSSWALK.items():
        val_e = ws.cell(row_idx, 5).value or 0.0
        val_f = ws.cell(row_idx, 6).value or 0.0
        for m in mappings:
            raw_amt = (val_e if m["source_col"] == "E" else val_f) or 0.0
            if isinstance(raw_amt, (int, float)) and raw_amt > 0:
                recipe = m["prep_recipe"]
                demands[recipe] = demands.get(recipe, 0.0) + (float(raw_amt) * m["multiplier"])
    return cat_date, demands

# ==============================================================================
# 3. STREAMLIT UI
# ==============================================================================
col1, col2 = st.columns(2)
with col1:
    prep_file = st.file_uploader("1. Upload Next-Day Prep Report (.xlsx)", type=["xlsx"])
with col2:
    cat_files = st.file_uploader("2. Upload ALL 3 Catering Calculator Files (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if prep_file and cat_files:
    wb_prep = openpyxl.load_workbook(prep_file)
    ws_prep = wb_prep['Prep Report']

    raw_prep_date = str(ws_prep.cell(4, 9).value)
    date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', raw_prep_date)
    prep_date = datetime.strptime(date_match.group(1), '%m/%d/%Y').date()
    st.info(f"**Prep Report Target Date (Day 1):** {prep_date.strftime('%A, %m/%d/%Y')}")

    # Map each uploaded file to its exact day offset (Day 1, Day 2, Day 3)
    day_demands = {1: {}, 2: {}, 3: {}}
    file_summary = []
    
    for f in cat_files:
        c_date, c_demands = extract_catering_demands(f.getvalue())
        day_offset = (c_date - prep_date).days + 1
        file_summary.append(f"📁 **{f.name}** -> Date: **{c_date.strftime('%m/%d/%Y')}** (Assigned to **Day {day_offset}**)")
        if day_offset in [1, 2, 3]:
            for rec, qty in c_demands.items():
                day_demands[day_offset][rec] = day_demands[day_offset].get(rec, 0.0) + qty

    st.success("\n".join(file_summary))

    # Update Prep Report Rows
    preview_rows = []
    for r in range(6, ws_prep.max_row + 1):
        rec_name = ws_prep.cell(r, 2).value
        sl_str = ws_prep.cell(r, 4).value
        bs_str = ws_prep.cell(r, 7).value
        iu_str = ws_prep.cell(r, 8).value
        orig_dc = ws_prep.cell(r, 11).value
        orig_slm = ws_prep.cell(r, 12).value

        if sl_str is not None and rec_name is not None:
            orig_dc_val = float(orig_dc or 0.0)
            orig_slm_val = float(orig_slm or 0.0)
            sl_days = parse_shelf_life_days(sl_str)
            batch_yield = parse_batch_size(bs_str, iu_str)

            # Day 1 Catering for Walk-In Trigger
            d1_cat = day_demands[1].get(rec_name, 0.0)
            # Multi-Day Catering capped by item's Shelf Life (1d, 2d, or 3d)
            window_cat = sum(day_demands[d].get(rec_name, 0.0) for d in range(1, sl_days + 1))
            window_cat_batches = window_cat / batch_yield if batch_yield > 0 else 0.0

            new_dc_val = round((orig_dc_val + d1_cat) * 4) / 4
            new_slm_val = round((orig_slm_val + window_cat_batches) * 4) / 4

            ws_prep.cell(r, 11).value = new_dc_val
            ws_prep.cell(r, 12).value = new_slm_val

            if d1_cat > 0 or window_cat > 0:
                preview_rows.append({
                    "Recipe Name": rec_name,
                    "Shelf Life": sl_str,
                    "Batch Size": bs_str,
                    "Inv. Unit": iu_str,
                    "Orig Daily Cons": orig_dc_val,
                    "+ Day 1 Catering": round(d1_cat, 2),
                    "NEW Daily Cons": new_dc_val,
                    "Orig Shelf Life Max (X)": orig_slm_val,
                    f"+ Catering Batches ({sl_days}d)": round(window_cat_batches, 2),
                    "NEW Shelf Life Max (X)": new_slm_val,
                })

    st.subheader("🔍 Recipes Updated by Multi-Day Catering")
    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True)

    output = BytesIO()
    wb_prep.save(output)
    st.download_button(
        label="📥 Download Updated Walk-In Prep Report (.xlsx)",
        data=output.getvalue(),
        file_name=f"Prep_Report_With_Catering_{prep_date.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
