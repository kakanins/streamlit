#CARA RUNNYA: streamlit run app.py --server.maxUploadSize=1024

import streamlit as st
import pandas as pd
from io import BytesIO
import math
import os

st.title("📊 Excel Filter")

uploaded_files = st.file_uploader(
    "Upload satu atau beberapa file Excel", type=["xlsx"], accept_multiple_files=True
)

if uploaded_files:
    combined_df = pd.DataFrame()

    for i, file in enumerate(uploaded_files):
        df = pd.read_excel(file, dtype=str)
        if i > 0:
            df = df.iloc[1:]
        combined_df = pd.concat([combined_df, df], ignore_index=True)

    st.success(f"{len(uploaded_files)} file berhasil digabung!")

    st.subheader("Data Preview")
    st.dataframe(combined_df.head())

    st.subheader("Filter Data")
    filter_columns = st.multiselect("Pilih kolom yang ingin difilter", combined_df.columns.tolist())

    filters = {}
    excludes = {}

    for col in filter_columns:
        unique_vals = combined_df[col].dropna().unique().tolist()
        selected_vals = st.multiselect(f"Pilih nilai untuk '{col}'", unique_vals, key=col)
        exclude = st.checkbox(f"❌ Kecualikan nilai ini dari kolom '{col}'", key=f"exclude_{col}")
        if selected_vals:
            if exclude:
                excludes[col] = selected_vals
            else:
                filters[col] = selected_vals

    st.subheader("Tambah / Ganti Kolom dengan Rumus Python")
    new_col_name = st.text_input("Nama kolom target:")
    formula = st.text_input("Rumus Python (misal: TOP - ANGS_AKH - 1):")

    with st.expander("➕ Buat Kolom Baru dengan Logika Kombinasi"):
        st.markdown("Gunakan builder ini untuk membuat aturan seperti: jika KOL_A adalah 'A' dan KOL_B adalah 'B', maka isi kolom = 'GOL 1'.")

        logic_col_name = st.text_input("Nama kolom hasil logika", value="KATEGORI")

        num_logic_rules = st.number_input("Jumlah aturan (if-else)", min_value=1, max_value=10, value=1, step=1)

        label_ops = {
            "Sama dengan": "==",
            "Tidak sama dengan": "!=",
            "Lebih besar dari": ">",
            "Lebih besar/sama": ">=",
            "Lebih kecil dari": "<",
            "Lebih kecil/sama": "<=",
            "Termasuk (daftar nilai)": "in",
            "Tidak termasuk (daftar nilai)": "not in",
            "Mengandung teks": "contains"
        }

        all_logic_rules = []
        for i in range(num_logic_rules):
            st.markdown(f"### Aturan {i+1}")
            n_conds = st.number_input(f"Jumlah kondisi dalam Aturan {i+1}", min_value=1, max_value=5, value=1, step=1, key=f"ncond_{i}")
            conds = []
            for j in range(n_conds):
                col = st.selectbox("Kolom", combined_df.columns.tolist(), key=f"lcol_{i}_{j}")
                label_op = st.selectbox("Operator", list(label_ops.keys()), key=f"lop_{i}_{j}")
                op_code = label_ops[label_op]

                if op_code in ["in", "not in"]:
                    unique_vals = combined_df[col].dropna().unique().tolist()
                    val = st.multiselect("Pilih nilai", unique_vals, key=f"lvalmulti_{i}_{j}")
                    val_str = ",".join(val)
                else:
                    val_str = st.text_input("Nilai (boleh pisahkan dengan koma jika lebih dari satu)", key=f"lval_{i}_{j}")

                conds.append((col, op_code, val_str))

            hasil = st.text_input(f"Isi kolom jika aturan {i+1} cocok", key=f"out_{i}")
            all_logic_rules.append((conds, hasil))

        default_val = st.text_input("Isi kolom jika tidak ada aturan yang cocok", value="LAINNYA")

    phone_filter = st.checkbox("📱 Filter baris dengan Nomor HP valid")

    if st.button("▶️ Proses Data"):
        filtered_df = combined_df.copy()

        for col, allowed_vals in filters.items():
            filtered_df = filtered_df[filtered_df[col].isin(allowed_vals)]

        for col, excluded_vals in excludes.items():
            filtered_df = filtered_df[~filtered_df[col].isin(excluded_vals)]

        if phone_filter:
            if all(col in filtered_df.columns for col in ['CUST_MOBPHONE', 'CUST_MOBPHONE_2']):
                def is_valid(row):
                    hp1 = str(row['CUST_MOBPHONE']) if pd.notna(row['CUST_MOBPHONE']) else ""
                    hp2 = str(row['CUST_MOBPHONE_2']) if pd.notna(row['CUST_MOBPHONE_2']) else ""
                    if hp1.strip() == "" and hp2.strip() == "":
                        return False
                    if not hp1.startswith("08") and not hp2.startswith("08"):
                        return False
                    return True
                filtered_df = filtered_df[filtered_df.apply(is_valid, axis=1)]
            else:
                st.warning("Kolom CUST_MOBPHONE dan CUST_MOBPHONE_2 tidak ditemukan di data.")

        if new_col_name and formula:
            try:
                temp_df = filtered_df.copy()
                temp_df = temp_df.apply(pd.to_numeric, errors='ignore').fillna(0)
                filtered_df[new_col_name] = temp_df.eval(formula)
                st.success(f"Kolom '{new_col_name}' berhasil ditambahkan/diperbarui.")
            except Exception as e:
                st.error(f"Gagal evaluasi rumus: {e}")

        def make_logic_code(rules, default_output):
            logic_lines = []
            for idx, (conds, output_val) in enumerate(rules):
                cond_strs = []
                for col, op, val in conds:
                    if op == "contains":
                        cond_str = f"'{val.lower()}' in str(row['{col}']).lower()"
                    elif op in ["in", "not in"]:
                        val_list = [f"'{v.strip()}'" for v in val.split(',') if v.strip()]
                        cond_str = f"row['{col}'] {op} [{', '.join(val_list)}]"
                    else:
                        cond_str = f"row['{col}'] {op} '{val}'"
                    cond_strs.append(cond_str)
                full_condition = " and ".join(cond_strs)
                if idx == 0:
                    logic_lines.append(f"if {full_condition}:")
                else:
                    logic_lines.append(f"elif {full_condition}:")
                logic_lines.append(f"    result = '{output_val}'")
            logic_lines.append(f"else:")
            logic_lines.append(f"    result = '{default_output}'")
            return "\n".join(logic_lines)

        if logic_col_name and all_logic_rules:
            try:
                logic_code = make_logic_code(all_logic_rules, default_val)
                def apply_logic(row):
                    local_vars = {'row': row}
                    exec("result = None\n" + logic_code, {}, local_vars)
                    return local_vars['result']
                filtered_df[logic_col_name] = filtered_df.apply(apply_logic, axis=1)
                st.success(f"Kolom '{logic_col_name}' berhasil dibuat dari logika kombinasi.")
                with st.expander("📜 Lihat rumus Python hasil konversi"):
                    st.code(logic_code, language='python')
            except Exception as e:
                st.error(f"Gagal evaluasi logika kombinasi: {e}")

        st.subheader("Filtered Result")
        st.write(f"{len(filtered_df)} baris hasil akhir")
        st.dataframe(filtered_df.head(100))

        base_filename = os.path.splitext(uploaded_files[0].name)[0] if len(uploaded_files) == 1 else "gabungan"
        max_rows = 1000000
        num_parts = math.ceil(len(filtered_df) / max_rows)

        for i in range(num_parts):
            start = i * max_rows
            end = start + max_rows
            part_df = filtered_df.iloc[start:end]

            output = BytesIO()
            part_df.to_excel(output, index=False)
            output.seek(0)

            filename = (
                f"filtered__{base_filename}_part{i+1}.xlsx"
                if num_parts > 1
                else f"filtered__{base_filename}.xlsx"
            )

            st.download_button(
                label=f"📥 Download hasil (Part {i+1}) - {len(part_df)} baris",
                data=output,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
