import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(
    page_title="Gestión de Artículos — Cómodo",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>.block-container { padding-top: 1.5rem; }</style>
""", unsafe_allow_html=True)

MESES      = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
MESES_FULL = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
              7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

GDRIVE_FILE_ID = "1w2I5XaswfouEzS7qZXnPde7smNTmP1--"   # mismo archivo de ventas
PASSWORD       = "gerencia2025"


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_peso(v):
    return f"${v:,.0f}".replace(",", ".")

def fmt_compacto(v):
    if v >= 1_000_000: return f"${v/1_000_000:.1f}M"
    if v >= 1_000:     return f"${v/1_000:.0f}K"
    return fmt_peso(v)


# ── Descarga desde Google Drive ───────────────────────────────────────────────
@st.cache_data(show_spinner="Descargando archivo...")
def cargar_desde_url(file_id):
    import requests, re
    errores = []
    urls = [
        f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t",
        f"https://drive.google.com/uc?export=download&id={file_id}&confirm=t",
        f"https://drive.google.com/uc?export=download&id={file_id}",
    ]
    for url in urls:
        try:
            r = requests.get(url, allow_redirects=True, timeout=90)
            ct = r.headers.get("Content-Type", "")
            if "text/html" in ct and r.status_code == 200:
                for pat in [r'confirm=([0-9A-Za-z_\-]+)', r'"([0-9A-Za-z_\-]{6,})"']:
                    m = re.search(pat, r.text)
                    if m and len(m.group(1)) > 3:
                        r2 = requests.get(
                            f"https://drive.usercontent.google.com/download?id={file_id}&export=download&confirm=t&uuid={m.group(1)}",
                            allow_redirects=True, timeout=90)
                        if r2.status_code == 200 and len(r2.content) > 1000:
                            return r2.content
            if r.status_code == 200 and len(r.content) > 1000 and b"html" not in r.content[:100].lower():
                return r.content
            errores.append(f"{url[:60]}: status={r.status_code}")
        except Exception as e:
            errores.append(f"{url[:60]}: {e}")
    return errores


# ── Carga y parseo del Excel ──────────────────────────────────────────────────
@st.cache_data(show_spinner="Procesando datos...")
def cargar_datos(archivo):
    import io

    raw = bytes(archivo) if isinstance(archivo, (bytes, bytearray)) else archivo.read()
    xls = pd.ExcelFile(io.BytesIO(raw), engine="calamine")
    hojas = xls.sheet_names

    def _find_col(df, keywords):
        for c in df.columns:
            lc = c.lower().replace(" ", "").replace("_", "")
            if any(k in lc for k in keywords):
                return c
        return None

    def _norm_cod(series):
        return (series.astype(str).str.strip()
                .str.replace(r'\.0+$', '', regex=True).str.upper())

    # -- Ventas --
    df_v = xls.parse("Ventas")
    nombres_v = [
        "fecha","periodo","cod_cliente","cliente","cod_vendedor","vendedor",
        "cbte","t_cbte","pto_vta","n_cbte","cod_articulo","descripcion",
        "cantidad","facturacion","familia","rubro","subrubro","marca",
        "clasificacion","subclasificacion","localidad","provincia",
    ]
    df_v = df_v.iloc[:, :len(nombres_v)]
    df_v.columns = nombres_v
    df_v["fecha"] = pd.to_datetime(df_v["fecha"], errors="coerce")
    df_v = df_v.dropna(subset=["fecha", "facturacion"])
    df_v["año"] = df_v["fecha"].dt.year
    df_v["mes"] = df_v["fecha"].dt.month
    df_v["cod_articulo"] = df_v["cod_articulo"].astype(str).str.strip()

    # -- Base artículos (descripción limpia) --
    df_art = None
    for h in hojas:
        hn = h.lower().replace(" ", "").replace("_", "")
        if "base" in hn and "art" in hn:
            try:
                da = xls.parse(h)
                da.columns = [str(c).strip() for c in da.columns]
                cc = _find_col(da, ["cod", "código", "codigo", "articulo", "art"])
                cd = _find_col(da, ["desc", "nombre", "detalle"])
                if cc and cd:
                    da = da.rename(columns={cc: "cod_articulo", cd: "descripcion"})
                    da["cod_articulo"] = _norm_cod(da["cod_articulo"])
                    da["descripcion"]  = da["descripcion"].astype(str).str.strip()
                    for campo, palabras in [
                        ("marca",    ["marca","brand"]),
                        ("familia",  ["familia","family","categ"]),
                        ("rubro",    ["rubro","rubr","tipo"]),
                        ("subrubro", ["subrubro","sub"]),
                    ]:
                        cx = _find_col(da, palabras)
                        if cx and cx not in ("cod_articulo", "descripcion"):
                            da = da.rename(columns={cx: campo})
                    cols_keep = [c for c in ["cod_articulo","descripcion","marca","familia","rubro","subrubro"]
                                 if c in da.columns]
                    df_art = da[cols_keep].dropna(subset=["cod_articulo"]).drop_duplicates("cod_articulo").copy()
            except Exception:
                pass
            break

    # -- Stock --
    df_stock = None
    for h in hojas:
        hn = h.lower()
        if df_stock is None and ("stock" in hn or "existencia" in hn):
            try:
                ds = xls.parse(h)
                ds.columns = [str(c).strip() for c in ds.columns]
                cc = _find_col(ds, ["cod", "código", "codigo", "articulo", "art"])
                cs = _find_col(ds, ["stock", "cantidad", "existencia", "saldo", "unid"])
                if cc and cs:
                    ds = ds.rename(columns={cc: "cod_articulo", cs: "stock_actual"})
                    ds["cod_articulo"] = _norm_cod(ds["cod_articulo"])
                    ds["stock_actual"] = pd.to_numeric(ds["stock_actual"], errors="coerce").fillna(0)
                    df_stock = ds[["cod_articulo","stock_actual"]].dropna(subset=["cod_articulo"]).copy()
            except Exception:
                pass

    # -- Precios --
    df_precios = None
    for h in hojas:
        hn = h.lower()
        if df_precios is None and any(k in hn for k in ["precio","lista","costo"]):
            try:
                dp = xls.parse(h)
                dp.columns = [str(c).strip() for c in dp.columns]
                cc = _find_col(dp, ["cod", "código", "codigo", "articulo", "art"])
                cp = _find_col(dp, ["precio", "costo", "lista", "unitario", "valor"])
                if cc and cp:
                    dp = dp.rename(columns={cc: "cod_articulo", cp: "precio_unitario"})
                    dp["cod_articulo"]    = _norm_cod(dp["cod_articulo"])
                    dp["precio_unitario"] = pd.to_numeric(dp["precio_unitario"], errors="coerce").fillna(0)
                    df_precios = dp[["cod_articulo","precio_unitario"]].dropna(subset=["cod_articulo"]).copy()
            except Exception:
                pass

    return df_v, df_art, df_stock, df_precios


# ── Widget principal de artículos ─────────────────────────────────────────────
def vista_articulos(ventas_df, df_art=None, df_stock=None, df_precios=None):

    hoy_art = ventas_df["fecha"].max()

    # Mapas de precio y stock
    precios_map = {}
    if df_precios is not None and not df_precios.empty:
        precios_map = df_precios.set_index("cod_articulo")["precio_unitario"].to_dict()
    stock_map = {}
    if df_stock is not None and not df_stock.empty:
        stock_map = df_stock.set_index("cod_articulo")["stock_actual"].to_dict()

    vdf = ventas_df.copy()
    vdf["cod_str"] = (vdf["cod_articulo"].astype(str).str.strip()
                      .str.replace(r'\.0+$', '', regex=True).str.upper())

    # Eje de tiempo
    time_df = (
        vdf[["año","mes"]].drop_duplicates()
        .sort_values(["año","mes"]).reset_index(drop=True)
    )
    time_df["mes_col"] = time_df.apply(
        lambda r: f"{MESES[r['mes']-1]}-{str(r['año'])[-2:]}", axis=1)
    month_cols = time_df["mes_col"].tolist()
    años_data  = sorted(time_df["año"].unique().tolist())
    vdf = vdf.merge(time_df[["año","mes","mes_col"]], on=["año","mes"], how="left")

    # Pivot unidades
    pivot = (
        vdf.groupby(["cod_str","mes_col"])["cantidad"]
        .sum().unstack("mes_col", fill_value=0)
        .reindex(columns=month_cols, fill_value=0)
    )

    # Totales por año
    for y in años_data:
        cols_y = [c for c in month_cols if c.endswith(f"-{str(y)[-2:]}")]
        pivot[f"__tot_{y}"] = pivot[cols_y].sum(axis=1)

    # Variación YoY
    if len(años_data) >= 2:
        y1, y0 = años_data[-1], años_data[-2]
        denom = pivot[f"__tot_{y0}"].where(pivot[f"__tot_{y0}"] > 0, other=pd.NA)
        pivot["__var_yoy"] = ((pivot[f"__tot_{y1}"] - pivot[f"__tot_{y0}"]) / denom * 100).round(1)
    else:
        pivot["__var_yoy"] = pd.NA

    # Promedios
    hoy_ts = pd.Timestamp(hoy_art.year, hoy_art.month, 1)
    for n, col in [(12,"__p12"), (6,"__p6"), (3,"__p3")]:
        cutoff = hoy_ts - pd.DateOffset(months=n)
        cols_n = [r.mes_col for r in time_df.itertuples()
                  if pd.Timestamp(r.año, r.mes, 1) >= cutoff]
        valid  = [c for c in cols_n if c in pivot.columns]
        pivot[col] = (pivot[valid].sum(axis=1) / n).round(1) if valid else 0.0
    pivot["__sin_mov"] = pivot["__p6"] == 0

    # Metadata desde ventas
    meta = (
        vdf.groupby("cod_str")
        .agg(
            cod_articulo=("cod_articulo","first"),
            descripcion =("descripcion", "first"),
            marca       =("marca",       "first"),
            familia     =("familia",     "first"),
            rubro       =("rubro",       "first"),
            subrubro    =("subrubro",    "first"),
        )
        .reset_index()
    )

    # Reemplazar con base artículos si existe
    if df_art is not None and not df_art.empty:
        for campo in ["descripcion","marca","familia","rubro","subrubro"]:
            if campo in df_art.columns:
                campo_map = df_art.set_index("cod_articulo")[campo].to_dict()
                meta[campo] = meta["cod_str"].map(campo_map).fillna(meta[campo])

    meta["stock"]     = meta["cod_str"].map(stock_map).fillna(0)
    meta["precio"]    = meta["cod_str"].map(precios_map).fillna(0)
    meta["val_stock"] = (meta["stock"] * meta["precio"]).round(0)

    # Combinar
    grp = meta.merge(pivot.reset_index(), on="cod_str", how="left")
    for mc in month_cols:
        if mc not in grp.columns: grp[mc] = 0
    for c in [f"__tot_{y}" for y in años_data] + ["__p12","__p6","__p3","__sin_mov","__var_yoy"]:
        if c not in grp.columns: grp[c] = 0
    grp["__sin_mov"] = grp["__sin_mov"].fillna(True).astype(bool)
    grp["__meses_stk"] = grp.apply(
        lambda r: round(r["stock"] / r["__p6"], 1) if r["__p6"] > 0
                  else ("Inf" if r["stock"] > 0 else None), axis=1)

    # ── Filtros ───────────────────────────────────────────────────────────────
    with st.expander("🔽 Filtros", expanded=True):
        fc1, fc2 = st.columns(2)
        busq_cod  = fc1.text_input("🔍 Código:", placeholder="Ej: TDCR...", key="bco")
        busq_desc = fc2.text_input("🔍 Descripción:", placeholder="Ej: AMOLADORA...", key="bde")
        f1, f2, f3, f4 = st.columns(4)
        sel_marca = f1.multiselect("Marca",    sorted(grp["marca"].dropna().unique()),    key="fm", placeholder="Todas")
        sel_fam   = f2.multiselect("Familia",  sorted(grp["familia"].dropna().unique()),  key="ff", placeholder="Todas")
        sel_rub   = f3.multiselect("Rubro",    sorted(grp["rubro"].dropna().unique()),    key="fr", placeholder="Todos")
        sel_sub   = f4.multiselect("Subrubro", sorted(grp["subrubro"].dropna().unique()), key="fs", placeholder="Todos")
        t1, t2 = st.columns(2)
        solo_stock = t1.toggle("📦 Solo con stock",           key="ts")
        solo_mov   = t2.toggle("📈 Solo con mov. últimos 6m", key="tm")

    if busq_cod.strip():
        grp = grp[grp["cod_str"].str.contains(busq_cod.strip().upper(), na=False)]
    if busq_desc.strip():
        grp = grp[grp["descripcion"].str.upper().str.contains(busq_desc.strip().upper(), na=False)]
    if sel_marca: grp = grp[grp["marca"].isin(sel_marca)]
    if sel_fam:   grp = grp[grp["familia"].isin(sel_fam)]
    if sel_rub:   grp = grp[grp["rubro"].isin(sel_rub)]
    if sel_sub:   grp = grp[grp["subrubro"].isin(sel_sub)]
    if solo_stock: grp = grp[grp["stock"] > 0]
    if solo_mov:   grp = grp[~grp["__sin_mov"]]

    if grp.empty:
        st.warning("No hay artículos para los filtros seleccionados.")
        return

    # ── KPIs ──────────────────────────────────────────────────────────────────
    ck = st.columns(4)
    ck[0].metric("📦 Artículos",   f"{len(grp):,}")
    ck[1].metric("📊 Uds totales", f"{sum(grp[f'__tot_{y}'].sum() for y in años_data):,.0f}")
    ck[2].metric("💵 Valor stock",  fmt_peso(grp["val_stock"].sum()))
    ck[3].metric("⚠️ Sin mov 6m",  f"{grp['__sin_mov'].sum():,}")

    st.markdown("---")

    # ── Ordenamiento ──────────────────────────────────────────────────────────
    ord_map = {f"Total {y}": f"__tot_{y}" for y in reversed(años_data)}
    ord_map.update({"Prom 6m": "__p6", "Stock": "stock", "Valor stock": "val_stock"})
    ord_sel = st.radio("Ordenar por:", list(ord_map.keys()), horizontal=True, key="ord")
    grp = grp.sort_values(ord_map[ord_sel], ascending=False, na_position="last").reset_index(drop=True)
    grp.insert(0, "#", range(1, len(grp)+1))

    # ── Tabla ─────────────────────────────────────────────────────────────────
    tbl = grp[["#","cod_articulo","descripcion","marca","familia","rubro","subrubro"]].rename(columns={
        "cod_articulo":"Código","descripcion":"Descripción",
        "marca":"Marca","familia":"Familia","rubro":"Rubro","subrubro":"Subrubro"
    }).copy()

    for mc in month_cols:
        tbl[mc] = grp[mc].fillna(0).astype(int)
    for y in años_data:
        tbl[f"Total {y}"] = grp[f"__tot_{y}"].fillna(0).astype(int)
    if len(años_data) >= 2:
        y1, y0 = años_data[-1], años_data[-2]
        tbl[f"{y1} vs {y0} (%)"] = grp["__var_yoy"]
    tbl["Prom 12m"] = grp["__p12"]
    tbl["Prom 6m"]  = grp["__p6"]
    tbl["Prom 3m"]  = grp["__p3"]
    if df_stock is not None:
        tbl["Stock"]       = grp["stock"].fillna(0).astype(int)
        tbl["Meses stock"] = grp["__meses_stk"].apply(
            lambda x: "—" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x))
    if precios_map:
        tbl["Valor stock ($)"] = grp["val_stock"].fillna(0)
    tbl["SIN MOV 6m"] = grp["__sin_mov"].apply(lambda x: "⚠️" if x else "")

    col_cfg = {}
    for mc in month_cols:
        col_cfg[mc] = st.column_config.NumberColumn(mc, format="%d", width="small")
    for y in años_data:
        col_cfg[f"Total {y}"] = st.column_config.NumberColumn(f"Total {y}", format="%d")
    if len(años_data) >= 2:
        col_cfg[f"{y1} vs {y0} (%)"] = st.column_config.NumberColumn(f"{y1} vs {y0} (%)", format="%.1f%%")
    col_cfg["Prom 12m"] = st.column_config.NumberColumn("Prom 12m", format="%.1f")
    col_cfg["Prom 6m"]  = st.column_config.NumberColumn("Prom 6m",  format="%.1f")
    col_cfg["Prom 3m"]  = st.column_config.NumberColumn("Prom 3m",  format="%.1f")
    if df_stock is not None:
        col_cfg["Stock"] = st.column_config.NumberColumn("Stock (22/5)", format="%d")
    if precios_map:
        col_cfg["Valor stock ($)"] = st.column_config.NumberColumn("Valor stock ($)", format="$ %,.0f")

    sel_grid = st.dataframe(
        tbl, use_container_width=True, hide_index=True, height=520,
        on_select="rerun", selection_mode="single-row",
        column_config=col_cfg,
    )
    st.download_button(
        "📥 Descargar tabla",
        grp.to_csv(index=False).encode("utf-8"),
        file_name="articulos_comodo.csv",
        mime="text/csv",
        key="dl_art",
    )

    # ── Composición valorizada ─────────────────────────────────────────────────
    st.markdown("---")
    with st.expander("💰 Composición valorizada por Marca / Familia / Rubro", expanded=False):
        cv1, cv2, cv3 = st.columns([2,1,1])
        comp_by   = cv1.radio("Agrupar por:", ["Marca","Familia","Rubro"], horizontal=True, key="cb")
        años_comp = ["Todos"] + [str(y) for y in años_data]
        sel_año_c = cv2.selectbox("Año:", años_comp, key="cay")
        meses_comp = ["Todos"] + [MESES_FULL[i] for i in range(1,13)]
        sel_mes_c  = cv3.selectbox("Mes:", meses_comp, key="cme")

        vf_comp = vdf.copy()
        if sel_año_c != "Todos":
            vf_comp = vf_comp[vf_comp["año"] == int(sel_año_c)]
        if sel_mes_c != "Todos":
            mes_num = next(k for k, v in MESES_FULL.items() if v == sel_mes_c)
            vf_comp = vf_comp[vf_comp["mes"] == mes_num]
        if sel_marca: vf_comp = vf_comp[vf_comp["marca"].isin(sel_marca)]
        if sel_fam:   vf_comp = vf_comp[vf_comp["familia"].isin(sel_fam)]
        if sel_rub:   vf_comp = vf_comp[vf_comp["rubro"].isin(sel_rub)]

        if not vf_comp.empty:
            vf_comp = vf_comp.copy()
            vf_comp["valor"] = vf_comp["cantidad"] * vf_comp["cod_str"].map(precios_map).fillna(0)
            usar_valor = bool(precios_map) and vf_comp["valor"].sum() > 0
            col_val   = "valor" if usar_valor else "facturacion"
            label_val = "Valor a precio lista ($)" if usar_valor else "Facturación ($)"
            col_grp_c = comp_by.lower()
            comp = (
                vf_comp.groupby(col_grp_c)[col_val].sum()
                .reset_index(name="total").dropna(subset=[col_grp_c])
                .pipe(lambda d: d[d["total"] > 0])
                .sort_values("total", ascending=False)
            )
            total_c = comp["total"].sum()
            comp["pct"] = comp["total"] / total_c * 100
            cg1, cg2 = st.columns([2,1])
            with cg1:
                fig_c = px.bar(comp.head(20), x="pct", y=col_grp_c, orientation="h",
                               title=f"Composición por {comp_by}",
                               labels={"pct":"% del total", col_grp_c:""},
                               text=comp.head(20)["pct"].apply(lambda x: f"{x:.1f}%"),
                               color_discrete_sequence=["#0066cc"])
                fig_c.update_traces(textposition="outside")
                fig_c.update_layout(yaxis={"categoryorder":"total ascending"}, margin=dict(r=100))
                st.plotly_chart(fig_c, use_container_width=True)
            with cg2:
                tbl_c = comp[[col_grp_c,"total","pct"]].copy()
                tbl_c["total"] = tbl_c["total"].apply(fmt_peso)
                tbl_c["pct"]   = tbl_c["pct"].apply(lambda x: f"{x:.1f}%")
                tbl_c.columns  = [comp_by, label_val, "% del total"]
                st.dataframe(tbl_c, use_container_width=True, hide_index=True)

    # ── Detalle de artículo ────────────────────────────────────────────────────
    st.markdown("---")
    st.markdown("#### 🔍 Detalle de artículo")

    def _limpiar():
        st.session_state["det_busq"] = ""
        st.session_state["art_sel"]  = None

    if sel_grid.selection.rows:
        idx = sel_grid.selection.rows[0]
        if idx < len(grp):
            st.session_state["art_sel"] = grp.iloc[idx]["cod_str"]

    col_sb, col_cl = st.columns([5, 1])
    with col_cl:
        st.button("🗑️ Limpiar", key="det_limpiar", on_click=_limpiar, use_container_width=True)
    with col_sb:
        busq_det = st.text_input(
            "Buscar:", placeholder="Ej: DISCO o TDCROA1151",
            key="det_busq", label_visibility="collapsed")

    if busq_det.strip():
        bl = busq_det.strip().upper()
        mask = (grp["descripcion"].str.upper().str.contains(bl, na=False) |
                grp["cod_str"].str.contains(bl, na=False))
        cands = grp[mask].reset_index(drop=True)
        if cands.empty:
            st.warning(f"No se encontró '{busq_det}'.")
            st.session_state["art_sel"] = None
        else:
            opts = cands.apply(lambda r: f"{r['cod_articulo']} — {r['descripcion']}", axis=1).tolist()
            if len(cands) > 1:
                sel_d = st.selectbox("Seleccioná:", range(len(cands)),
                                     format_func=lambda i: opts[i],
                                     key=f"det_cand_{bl[:12]}")
                st.session_state["art_sel"] = cands.iloc[sel_d]["cod_str"]
            else:
                st.session_state["art_sel"] = cands.iloc[0]["cod_str"]
    elif not sel_grid.selection.rows:
        st.info("Hacé clic en un artículo de la tabla o buscá por nombre/código.")

    cod_det = st.session_state.get("art_sel")
    if cod_det is None:
        return

    art_r = meta[meta["cod_str"] == cod_det]
    if art_r.empty:
        return
    r = art_r.iloc[0]
    gr = grp[grp["cod_str"] == cod_det]
    if gr.empty:
        return
    gr = gr.iloc[0]

    st.markdown(f"### {r['cod_articulo']} — {r['descripcion']}")
    da_c = st.columns(4)
    tot_hist = sum(gr.get(f"__tot_{y}", 0) for y in años_data)
    da_c[0].metric("Uds históricas",    f"{tot_hist:,.0f}")
    da_c[1].metric("Prom mensual 6m",   f"{gr['__p6']:.1f}")
    if r["stock"] > 0:
        da_c[2].metric("Stock actual", f"{int(r['stock']):,}")
    if gr["__meses_stk"] is not None:
        da_c[3].metric("Meses de stock", str(gr["__meses_stk"]))

    vf_art = ventas_df[ventas_df["cod_articulo"].astype(str).str.strip()
                        .str.replace(r'\.0+$','',regex=True).str.upper() == cod_det].copy()
    vf_art["periodo"] = pd.to_datetime(
        vf_art["año"].astype(str) + "-" + vf_art["mes"].astype(str).str.zfill(2) + "-01")
    evol = vf_art.groupby("periodo").agg(
        unidades   =("cantidad",    "sum"),
        facturacion=("facturacion", "sum"),
        clientes   =("cod_cliente", "nunique"),
    ).reset_index().sort_values("periodo")

    prom_u = evol["unidades"].mean()
    prom_f = evol["facturacion"].mean()
    g1, g2 = st.columns(2)
    with g1:
        fig_u = px.bar(evol, x="periodo", y="unidades",
                       title="Unidades por mes (histórico)",
                       labels={"periodo":"","unidades":"Unidades"},
                       color_discrete_sequence=["#0066cc"])
        fig_u.add_hline(y=prom_u, line_dash="dash", line_color="orange",
                        annotation_text=f"Prom: {prom_u:,.1f}")
        fig_u.update_layout(xaxis_tickformat="%b %Y")
        st.plotly_chart(fig_u, use_container_width=True)
    with g2:
        fig_f = px.bar(evol, x="periodo", y="facturacion",
                       title="Facturación por mes (histórico)",
                       labels={"periodo":"","facturacion":"Facturación ($)"},
                       color_discrete_sequence=["#28a745"])
        fig_f.add_hline(y=prom_f, line_dash="dash", line_color="orange",
                        annotation_text=f"Prom: {fmt_compacto(prom_f)}")
        fig_f.update_layout(xaxis_tickformat="%b %Y")
        st.plotly_chart(fig_f, use_container_width=True)

    tbl_evol = evol.copy()
    tbl_evol["periodo"]     = tbl_evol["periodo"].dt.strftime("%b %Y")
    tbl_evol["unidades"]    = tbl_evol["unidades"].apply(lambda x: f"{x:,.0f}")
    tbl_evol["facturacion"] = tbl_evol["facturacion"].apply(fmt_peso)
    tbl_evol["clientes"]    = tbl_evol["clientes"].astype(int)
    tbl_evol.columns        = ["Período","Unidades","Facturación","Clientes"]
    st.dataframe(tbl_evol, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### 👥 Clientes que más compran este artículo")
    if not vf_art.empty:
        crit = st.radio("Ordenar por:", ["Unidades","Facturación"], horizontal=True, key="cli_crit")
        col_ord = "unidades" if crit == "Unidades" else "facturacion"
        top_cli = (
            vf_art.groupby("cod_cliente").agg(
                unidades     =("cantidad",    "sum"),
                facturacion  =("facturacion", "sum"),
                ultima_compra=("fecha",       "max"),
                cliente      =("cliente",     "first"),
            ).reset_index()
            .sort_values(col_ord, ascending=False).reset_index(drop=True)
        )
        top_cli.insert(0, "#", range(1, len(top_cli)+1))
        top10 = top_cli.head(10)
        fig_cli = px.bar(top10, x=col_ord, y="cliente", orientation="h",
                         title="Top 10 clientes",
                         labels={col_ord: crit, "cliente":""},
                         color_discrete_sequence=["#0066cc"],
                         text=top10[col_ord].apply(
                             lambda x: f"{x:,.0f}" if crit=="Unidades" else fmt_compacto(x)))
        fig_cli.update_traces(textposition="outside")
        fig_cli.update_layout(yaxis={"categoryorder":"total ascending"}, margin=dict(t=40,r=100))
        st.plotly_chart(fig_cli, use_container_width=True)

        tbl_cli = top_cli[["#","cliente","unidades","facturacion","ultima_compra"]].copy()
        tbl_cli["unidades"]      = tbl_cli["unidades"].apply(lambda x: f"{x:,.0f}")
        tbl_cli["facturacion"]   = tbl_cli["facturacion"].apply(fmt_peso)
        tbl_cli["ultima_compra"] = tbl_cli["ultima_compra"].dt.strftime("%d/%m/%Y")
        tbl_cli.columns = ["#","Cliente","Unidades","Facturación","Última compra"]
        st.dataframe(tbl_cli, use_container_width=True, hide_index=True)
        st.download_button(
            "📥 Descargar ranking de clientes",
            tbl_cli.to_csv(index=False).encode("utf-8"),
            file_name=f"clientes_{r['cod_articulo']}.csv",
            mime="text/csv", key="cli_dl"
        )


# ── App principal ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📦 Gestión de Artículos")
    st.markdown("Distribuidora Cómodo")
    st.markdown("---")
    pwd = st.text_input("Contraseña:", type="password")
    st.markdown("---")
    if st.button("🔄 Actualizar datos", use_container_width=True,
                 help="Recarga el archivo desde Google Drive"):
        st.cache_data.clear()
        st.rerun()

if pwd != PASSWORD:
    st.title("📦 Gestión de Artículos — Cómodo")
    st.info("🔐 Ingresá la contraseña en el panel izquierdo para acceder.")
    st.stop()

archivo = cargar_desde_url(GDRIVE_FILE_ID)

if archivo is None or isinstance(archivo, list):
    st.error("⚠️ No se pudo descargar el archivo de ventas desde Google Drive.")
    if isinstance(archivo, list):
        with st.expander("Detalle del error"):
            for e in archivo:
                st.code(e)
    if st.button("🔄 Reintentar"):
        st.cache_data.clear()
        st.rerun()
    st.stop()

df_ventas, df_art, df_stock, df_precios = cargar_datos(archivo)
hoy = df_ventas["fecha"].max()

st.title("📦 Gestión de Artículos")
st.caption(f"Datos al {hoy.strftime('%d/%m/%Y')}")
st.divider()

vista_articulos(df_ventas, df_art=df_art, df_stock=df_stock, df_precios=df_precios)
