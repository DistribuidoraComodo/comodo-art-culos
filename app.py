import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(
    page_title="Gestión de Artículos — Cómodo",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container { padding-top: 1.5rem; }
div[data-testid="metric-container"] {
    background: #f8f9fa;
    border: 1px solid #e0e0e0;
    border-radius: 8px;
    padding: 12px 16px;
}
</style>
""", unsafe_allow_html=True)

MESES      = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
MESES_FULL = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
              7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

GDRIVE_FILE_ID = "1w2I5XaswfouEzS7qZXnPde7smNTmP1--"
PASSWORD       = "gerencia2025"


# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt_peso(v):
    try:    return f"${v:,.0f}".replace(",", ".")
    except: return "$0"

def fmt_compacto(v):
    try:
        if abs(v) >= 1_000_000: return f"${v/1_000_000:.1f}M"
        if abs(v) >= 1_000:     return f"${v/1_000:.0f}K"
        return fmt_peso(v)
    except: return "$0"

def _norm_cod(series):
    return series.astype(str).str.strip().str.replace(r'\.0+$', '', regex=True).str.upper()


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


# ── Vista principal ───────────────────────────────────────────────────────────
def vista_articulos(ventas_df, df_art=None, df_stock=None, df_precios=None):

    hoy_art = ventas_df["fecha"].max()

    precios_map = {}
    if df_precios is not None and not df_precios.empty:
        precios_map = df_precios.set_index("cod_articulo")["precio_unitario"].to_dict()
    stock_map = {}
    if df_stock is not None and not df_stock.empty:
        stock_map = df_stock.set_index("cod_articulo")["stock_actual"].to_dict()

    vdf = ventas_df.copy()
    vdf["cod_str"] = _norm_cod(vdf["cod_articulo"])

    time_df = (
        vdf[["año","mes"]].drop_duplicates()
        .sort_values(["año","mes"]).reset_index(drop=True)
    )
    time_df["mes_col"] = time_df.apply(
        lambda r: f"{MESES[r['mes']-1]}-{str(r['año'])[-2:]}", axis=1)
    month_cols = time_df["mes_col"].tolist()
    años_data  = sorted(time_df["año"].unique().tolist())
    vdf = vdf.merge(time_df[["año","mes","mes_col"]], on=["año","mes"], how="left")

    pivot = (
        vdf.groupby(["cod_str","mes_col"])["cantidad"]
        .sum().unstack("mes_col", fill_value=0)
        .reindex(columns=month_cols, fill_value=0)
    )

    for y in años_data:
        cols_y = [c for c in month_cols if c.endswith(f"-{str(y)[-2:]}")]
        pivot[f"__tot_{y}"] = pivot[cols_y].sum(axis=1)

    if len(años_data) >= 2:
        y1, y0 = años_data[-1], años_data[-2]
        denom = pivot[f"__tot_{y0}"].where(pivot[f"__tot_{y0}"] > 0, other=pd.NA)
        pivot["__var_yoy"] = ((pivot[f"__tot_{y1}"] - pivot[f"__tot_{y0}"]) / denom * 100).round(1)
    else:
        pivot["__var_yoy"] = pd.NA

    hoy_ts = pd.Timestamp(hoy_art.year, hoy_art.month, 1)
    for n, col in [(12,"__p12"), (6,"__p6"), (3,"__p3")]:
        cutoff = hoy_ts - pd.DateOffset(months=n)
        cols_n = [r.mes_col for r in time_df.itertuples()
                  if pd.Timestamp(r.año, r.mes, 1) >= cutoff]
        valid  = [c for c in cols_n if c in pivot.columns]
        pivot[col] = (pivot[valid].sum(axis=1) / n).round(1) if valid else 0.0
    pivot["__sin_mov"] = pivot["__p6"] == 0

    cutoff_12 = hoy_ts - pd.DateOffset(months=12)
    cols_12 = [r.mes_col for r in time_df.itertuples()
               if pd.Timestamp(r.año, r.mes, 1) >= cutoff_12]
    valid_12 = [c for c in cols_12 if c in pivot.columns]
    pivot["__tot_12m"] = pivot[valid_12].sum(axis=1) if valid_12 else 0

    meta = (
        vdf.groupby("cod_str")
        .agg(
            cod_articulo=("cod_articulo","first"),
            descripcion =("descripcion", "first"),
            marca       =("marca",       "first"),
            familia     =("familia",     "first"),
            rubro       =("rubro",       "first"),
            subrubro    =("subrubro",    "first"),
        ).reset_index()
    )

    if df_art is not None and not df_art.empty:
        for campo in ["descripcion","marca","familia","rubro","subrubro"]:
            if campo in df_art.columns:
                campo_map = df_art.set_index("cod_articulo")[campo].to_dict()
                meta[campo] = meta["cod_str"].map(campo_map).fillna(meta[campo])

    meta["stock"]     = meta["cod_str"].map(stock_map).fillna(0)
    meta["precio"]    = meta["cod_str"].map(precios_map).fillna(0)
    meta["val_stock"] = (meta["stock"] * meta["precio"]).round(0)

    grp_full = meta.merge(pivot.reset_index(), on="cod_str", how="left")
    for mc in month_cols:
        if mc not in grp_full.columns: grp_full[mc] = 0
    for c in [f"__tot_{y}" for y in años_data] + ["__p12","__p6","__p3","__sin_mov","__var_yoy","__tot_12m"]:
        if c not in grp_full.columns: grp_full[c] = 0
    grp_full["__sin_mov"] = grp_full["__sin_mov"].fillna(True).astype(bool)
    grp_full["__meses_stk"] = grp_full.apply(
        lambda r: round(r["stock"] / r["__p3"], 1) if r["__p3"] > 0
                  else ("∞" if r["stock"] > 0 else None), axis=1)
    grp_full["__riesgo"] = (
        (grp_full["stock"] < grp_full["__p3"] * 2) &
        (grp_full["__p3"] > 0)
    )

    # ─────────────────────────────────────────────────────────────────────────
    tab_res, tab_art = st.tabs(["📊 Resumen ejecutivo", "📋 Artículos"])

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 1 — RESUMEN EJECUTIVO
    # ══════════════════════════════════════════════════════════════════════════
    with tab_res:

        total_art        = len(grp_full)
        activos_12m      = int((grp_full["__tot_12m"] > 0).sum())
        sin_mov_6m       = int(grp_full["__sin_mov"].sum())
        val_stk_tot      = grp_full["val_stock"].sum()
        riesgo_q         = int(grp_full["__riesgo"].sum())
        sin_stk_activos  = int(((grp_full["stock"] == 0) & (grp_full["__p3"] > 0)).sum())

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("📦 Total artículos",   f"{total_art:,}")
        k2.metric("✅ Activos (12m)",      f"{activos_12m:,}",
                  help="Con al menos 1 venta en los últimos 12 meses")
        k3.metric("⚠️ Sin mov. 6m",        f"{sin_mov_6m:,}",
                  help="Sin ventas en los últimos 6 meses")
        k4.metric("💵 Stock valorizado",   fmt_compacto(val_stk_tot))
        k5.metric("🚨 Riesgo quiebre",     f"{riesgo_q:,}",
                  help="Stock menor a 2 meses de demanda (prom. 3m)")
        k6.metric("📭 Sin stock + activos", f"{sin_stk_activos:,}",
                  help="Stock = 0 con promedio de ventas > 0 en últimos 3 meses")

        st.markdown("---")

        # ── Gráficos ──────────────────────────────────────────────────────────
        ch1, ch2, ch3 = st.columns([2, 2, 1.5])

        # Tendencia últimos 12 meses
        with ch1:
            vdf_12 = vdf.copy()
            vdf_12["periodo_ts"] = pd.to_datetime(
                vdf_12["año"].astype(str) + "-" + vdf_12["mes"].astype(str).str.zfill(2) + "-01")
            vdf_12 = vdf_12[vdf_12["periodo_ts"] >= cutoff_12]
            tend = (vdf_12.groupby("periodo_ts")
                    .agg(unidades=("cantidad","sum"), facturacion=("facturacion","sum"))
                    .reset_index().sort_values("periodo_ts"))
            tend["mes_lbl"] = tend["periodo_ts"].dt.strftime("%b %Y")

            fig_tend = go.Figure()
            fig_tend.add_trace(go.Bar(
                x=tend["mes_lbl"], y=tend["unidades"],
                name="Unidades", marker_color="#0066cc", opacity=0.75))
            fig_tend.add_trace(go.Scatter(
                x=tend["mes_lbl"], y=tend["unidades"],
                mode="lines+markers", name="Tendencia",
                line=dict(color="#ff6600", width=2), marker=dict(size=5)))
            fig_tend.update_layout(
                title="📈 Ventas totales — últimos 12 meses",
                showlegend=False, height=320,
                margin=dict(l=0, r=0, t=40, b=0),
                xaxis_tickangle=-40)
            st.plotly_chart(fig_tend, use_container_width=True)

        # Top 10 artículos
        with ch2:
            top10 = grp_full.nlargest(10, "__tot_12m")[["descripcion","__tot_12m"]].copy()
            top10["desc_short"] = top10["descripcion"].str[:28]
            fig_top = px.bar(
                top10, x="__tot_12m", y="desc_short", orientation="h",
                title="🏆 Top 10 artículos — últimos 12 meses",
                labels={"__tot_12m": "Unidades", "desc_short": ""},
                color_discrete_sequence=["#0066cc"],
                text=top10["__tot_12m"].apply(lambda x: f"{x:,.0f}"))
            fig_top.update_traces(textposition="outside")
            fig_top.update_layout(
                yaxis={"categoryorder": "total ascending"},
                height=320, margin=dict(l=0, r=70, t=40, b=0))
            st.plotly_chart(fig_top, use_container_width=True)

        # Mix por marca
        with ch3:
            mix = (grp_full.groupby("marca")["__tot_12m"]
                   .sum().reset_index()
                   .pipe(lambda d: d[d["__tot_12m"] > 0])
                   .sort_values("__tot_12m", ascending=False))
            if not mix.empty:
                otros = mix.iloc[7:]
                mix_pie = mix.head(7)
                if len(otros) > 0:
                    mix_pie = pd.concat([
                        mix_pie,
                        pd.DataFrame({"marca": ["Otros"], "__tot_12m": [otros["__tot_12m"].sum()]})
                    ], ignore_index=True)
                fig_mix = px.pie(
                    mix_pie, names="marca", values="__tot_12m",
                    title="🏷️ Mix por marca (últ. 12m)", hole=0.38)
                fig_mix.update_layout(
                    height=320, margin=dict(l=0, r=0, t=40, b=0),
                    legend=dict(font=dict(size=9), orientation="v"))
                fig_mix.update_traces(textposition="inside", textinfo="percent+label",
                                      textfont_size=10)
                st.plotly_chart(fig_mix, use_container_width=True)

        st.markdown("---")

        # ── Alertas ───────────────────────────────────────────────────────────
        al1, al2 = st.columns(2)

        with al1:
            riesgo_df = grp_full[grp_full["__riesgo"]].sort_values("__p3", ascending=False)
            lbl_riesgo = f"🚨 Riesgo de quiebre — {len(riesgo_df)} artículos"
            with st.expander(lbl_riesgo, expanded=riesgo_q > 0):
                if riesgo_df.empty:
                    st.success("No hay artículos en riesgo de quiebre.")
                else:
                    r_tbl = riesgo_df[["cod_articulo","descripcion","marca","stock","__p3","__meses_stk"]].head(25).copy()
                    r_tbl["__meses_stk"] = r_tbl["__meses_stk"].apply(
                        lambda x: "—" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x))
                    r_tbl.columns = ["Código","Descripción","Marca","Stock","Prom 3m","Meses cob."]
                    st.dataframe(r_tbl, use_container_width=True, hide_index=True, height=300)
                    st.download_button(
                        "📥 Descargar", r_tbl.to_csv(index=False).encode("utf-8"),
                        file_name="riesgo_quiebre.csv", mime="text/csv", key="dl_riesgo")

        with al2:
            ss_df = grp_full[(grp_full["stock"] == 0) & (grp_full["__p3"] > 0)].sort_values("__p3", ascending=False)
            lbl_ss = f"📭 Sin stock pero con demanda — {len(ss_df)} artículos"
            with st.expander(lbl_ss, expanded=sin_stk_activos > 0):
                if ss_df.empty:
                    st.success("No hay artículos sin stock con demanda activa.")
                else:
                    ss_tbl = ss_df[["cod_articulo","descripcion","marca","__p3","__p6","__p12"]].head(25).copy()
                    ss_tbl.columns = ["Código","Descripción","Marca","Prom 3m","Prom 6m","Prom 12m"]
                    st.dataframe(ss_tbl, use_container_width=True, hide_index=True, height=300)
                    st.download_button(
                        "📥 Descargar", ss_tbl.to_csv(index=False).encode("utf-8"),
                        file_name="sin_stock_activos.csv", mime="text/csv", key="dl_ss")

        st.markdown("---")

        # ── Composición por clasificación ─────────────────────────────────────
        st.markdown("#### 📊 Composición de ventas por clasificación")
        cc1, cc2 = st.columns([3, 1])
        comp_by  = cc1.radio("Agrupar por:",
                             ["Marca","Familia","Rubro","Subrubro"],
                             horizontal=True, key="res_cb")
        año_comp = cc2.selectbox("Año:", ["Todos"] + [str(y) for y in reversed(años_data)], key="res_ay")

        vf_c = vdf.copy()
        if año_comp != "Todos":
            vf_c = vf_c[vf_c["año"] == int(año_comp)]

        comp_col = comp_by.lower()
        if comp_col in vf_c.columns:
            comp = (vf_c.groupby(comp_col)["cantidad"].sum()
                    .reset_index(name="total")
                    .dropna(subset=[comp_col])
                    .pipe(lambda d: d[d["total"] > 0])
                    .sort_values("total", ascending=False))
            total_c = comp["total"].sum()
            comp["pct"] = comp["total"] / total_c * 100
            cg1, cg2 = st.columns([3, 1])
            with cg1:
                fig_c = px.bar(
                    comp.head(20), x="pct", y=comp_col, orientation="h",
                    title=f"Composición por {comp_by} — {año_comp}",
                    labels={"pct": "% del total", comp_col: ""},
                    text=comp.head(20)["pct"].apply(lambda x: f"{x:.1f}%"),
                    color_discrete_sequence=["#0066cc"])
                fig_c.update_traces(textposition="outside")
                fig_c.update_layout(yaxis={"categoryorder": "total ascending"},
                                    margin=dict(r=80), height=400)
                st.plotly_chart(fig_c, use_container_width=True)
            with cg2:
                comp_tbl = comp[[comp_col,"total","pct"]].copy()
                comp_tbl["total"] = comp_tbl["total"].apply(lambda x: f"{x:,.0f}")
                comp_tbl["pct"]   = comp_tbl["pct"].apply(lambda x: f"{x:.1f}%")
                comp_tbl.columns  = [comp_by, "Unidades", "% total"]
                st.dataframe(comp_tbl, use_container_width=True, hide_index=True, height=400)

    # ══════════════════════════════════════════════════════════════════════════
    # TAB 2 — ARTÍCULOS (grilla + detalle)
    # ══════════════════════════════════════════════════════════════════════════
    with tab_art:

        # ── Filtros ───────────────────────────────────────────────────────────
        with st.expander("🔽 Filtros", expanded=True):
            fc1, fc2 = st.columns(2)
            busq_cod  = fc1.text_input("🔍 Código:", placeholder="Ej: TDCR...", key="bco")
            busq_desc = fc2.text_input("🔍 Descripción:", placeholder="Ej: AMOLADORA...", key="bde")
            f1, f2, f3, f4 = st.columns(4)
            sel_marca = f1.multiselect("Marca",    sorted(grp_full["marca"].dropna().unique()),    key="fm", placeholder="Todas")
            sel_fam   = f2.multiselect("Familia",  sorted(grp_full["familia"].dropna().unique()),  key="ff", placeholder="Todas")
            sel_rub   = f3.multiselect("Rubro",    sorted(grp_full["rubro"].dropna().unique()),    key="fr", placeholder="Todos")
            sel_sub   = f4.multiselect("Subrubro", sorted(grp_full["subrubro"].dropna().unique()), key="fs", placeholder="Todos")
            t1, t2, t3 = st.columns(3)
            solo_stock  = t1.toggle("📦 Solo con stock",          key="ts")
            solo_mov    = t2.toggle("📈 Solo con mov. 6m",        key="tm")
            solo_riesgo = t3.toggle("🚨 Solo riesgo de quiebre",  key="tr")

        grp = grp_full.copy()
        if busq_cod.strip():
            grp = grp[grp["cod_str"].str.contains(busq_cod.strip().upper(), na=False)]
        if busq_desc.strip():
            grp = grp[grp["descripcion"].str.upper().str.contains(busq_desc.strip().upper(), na=False)]
        if sel_marca: grp = grp[grp["marca"].isin(sel_marca)]
        if sel_fam:   grp = grp[grp["familia"].isin(sel_fam)]
        if sel_rub:   grp = grp[grp["rubro"].isin(sel_rub)]
        if sel_sub:   grp = grp[grp["subrubro"].isin(sel_sub)]
        if solo_stock:  grp = grp[grp["stock"] > 0]
        if solo_mov:    grp = grp[~grp["__sin_mov"]]
        if solo_riesgo: grp = grp[grp["__riesgo"]]

        if grp.empty:
            st.warning("No hay artículos para los filtros seleccionados.")
            return

        # ── KPIs filtrados ────────────────────────────────────────────────────
        fk1, fk2, fk3, fk4 = st.columns(4)
        fk1.metric("📦 Artículos (filtro)",  f"{len(grp):,}")
        fk2.metric("📊 Uds últimos 12m",     f"{grp['__tot_12m'].sum():,.0f}")
        fk3.metric("💵 Stock valorizado",     fmt_compacto(grp["val_stock"].sum()))
        fk4.metric("⚠️ Sin mov. 6m",          f"{grp['__sin_mov'].sum():,}")

        st.markdown("---")

        # ── Ordenamiento ──────────────────────────────────────────────────────
        ord_map = {"Últimos 12m": "__tot_12m"}
        ord_map.update({f"Total {y}": f"__tot_{y}" for y in reversed(años_data)})
        ord_map.update({"Prom 3m": "__p3", "Prom 6m": "__p6",
                        "Stock": "stock", "Valor stock": "val_stock"})
        ord_sel = st.radio("Ordenar por:", list(ord_map.keys()), horizontal=True, key="ord")
        grp = grp.sort_values(ord_map[ord_sel], ascending=False, na_position="last").reset_index(drop=True)
        grp.insert(0, "#", range(1, len(grp)+1))

        # ── Tabla ─────────────────────────────────────────────────────────────
        tbl = grp[["#","cod_articulo","descripcion","marca","familia","rubro","subrubro"]].rename(columns={
            "cod_articulo": "Código", "descripcion": "Descripción",
            "marca": "Marca", "familia": "Familia", "rubro": "Rubro", "subrubro": "Subrubro"
        }).copy()

        for mc in month_cols:
            tbl[mc] = grp[mc].fillna(0).astype(int)
        for y in años_data:
            tbl[f"Total {y}"] = grp[f"__tot_{y}"].fillna(0).astype(int)
        if len(años_data) >= 2:
            y1, y0 = años_data[-1], años_data[-2]
            tbl[f"Var {y1}/{y0} (%)"] = grp["__var_yoy"]
        tbl["Prom 12m"] = grp["__p12"]
        tbl["Prom 6m"]  = grp["__p6"]
        tbl["Prom 3m"]  = grp["__p3"]
        if df_stock is not None:
            tbl["Stock"]      = grp["stock"].fillna(0).astype(int)
            tbl["Cob. meses"] = grp["__meses_stk"].apply(
                lambda x: "—" if x is None or (isinstance(x, float) and pd.isna(x)) else str(x))
        if precios_map:
            tbl["Valor stock ($)"] = grp["val_stock"].fillna(0)
        tbl["Estado"] = grp.apply(
            lambda r: "🚨 Quiebre" if r["__riesgo"]
                      else ("⚠️ Sin mov" if r["__sin_mov"] else "✅"), axis=1)

        col_cfg = {}
        for mc in month_cols:
            col_cfg[mc] = st.column_config.NumberColumn(mc, format="%d", width="small")
        for y in años_data:
            col_cfg[f"Total {y}"] = st.column_config.NumberColumn(f"Total {y}", format="%d")
        if len(años_data) >= 2:
            col_cfg[f"Var {y1}/{y0} (%)"] = st.column_config.NumberColumn(
                f"Var {y1}/{y0} (%)", format="%.1f%%")
        col_cfg["Prom 12m"] = st.column_config.NumberColumn("Prom 12m", format="%.1f")
        col_cfg["Prom 6m"]  = st.column_config.NumberColumn("Prom 6m",  format="%.1f")
        col_cfg["Prom 3m"]  = st.column_config.NumberColumn("Prom 3m",  format="%.1f")
        if df_stock is not None:
            col_cfg["Stock"] = st.column_config.NumberColumn("Stock", format="%d")
        if precios_map:
            col_cfg["Valor stock ($)"] = st.column_config.NumberColumn("Valor stock ($)", format="$ %,.0f")

        sel_grid = st.dataframe(
            tbl, use_container_width=True, hide_index=True, height=500,
            on_select="rerun", selection_mode="single-row",
            column_config=col_cfg,
        )
        dl1, dl2 = st.columns([1, 5])
        dl1.download_button(
            "📥 Descargar tabla",
            grp.to_csv(index=False).encode("utf-8"),
            file_name="articulos_comodo.csv",
            mime="text/csv", key="dl_art",
        )

        # ── Detalle de artículo ────────────────────────────────────────────────
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

        # Header del artículo
        st.markdown(f"### {r['cod_articulo']} — {r['descripcion']}")
        badge_estado = "🚨 Riesgo de quiebre" if gr["__riesgo"] else ("⚠️ Sin movimiento 6m" if gr["__sin_mov"] else "✅ Activo")
        st.caption(f"**{r.get('marca','')}** | {r.get('familia','')} › {r.get('rubro','')} › {r.get('subrubro','')}   {badge_estado}")

        # KPIs del artículo
        da_c = st.columns(5)
        tot_hist = sum(gr.get(f"__tot_{y}", 0) for y in años_data)
        da_c[0].metric("Uds históricas",    f"{tot_hist:,.0f}")
        da_c[1].metric("Prom mensual 12m",  f"{gr['__p12']:.1f}")
        da_c[2].metric("Prom mensual 6m",   f"{gr['__p6']:.1f}")
        da_c[3].metric("Prom mensual 3m",   f"{gr['__p3']:.1f}")
        if df_stock is not None:
            stk_val = int(r["stock"])
            meses_c = gr["__meses_stk"]
            meses_lbl = str(meses_c) if meses_c is not None and not (isinstance(meses_c, float) and pd.isna(meses_c)) else "—"
            da_c[4].metric("Stock actual", f"{stk_val:,}", delta=f"{meses_lbl} meses de cobertura",
                           delta_color="normal" if meses_lbl not in ("—","0") else "off")

        # Gráficos de evolución
        vf_art = ventas_df[_norm_cod(ventas_df["cod_articulo"]) == cod_det].copy()
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
                           title="Unidades por mes",
                           labels={"periodo":"","unidades":"Unidades"},
                           color_discrete_sequence=["#0066cc"])
            fig_u.add_hline(y=prom_u, line_dash="dash", line_color="orange",
                            annotation_text=f"Prom: {prom_u:,.1f}")
            fig_u.update_layout(xaxis_tickformat="%b %Y", height=300,
                                margin=dict(l=0,r=0,t=40,b=0))
            st.plotly_chart(fig_u, use_container_width=True)
        with g2:
            fig_f = px.bar(evol, x="periodo", y="facturacion",
                           title="Facturación por mes",
                           labels={"periodo":"","facturacion":"Facturación ($)"},
                           color_discrete_sequence=["#28a745"])
            fig_f.add_hline(y=prom_f, line_dash="dash", line_color="orange",
                            annotation_text=f"Prom: {fmt_compacto(prom_f)}")
            fig_f.update_layout(xaxis_tickformat="%b %Y", height=300,
                                margin=dict(l=0,r=0,t=40,b=0))
            st.plotly_chart(fig_f, use_container_width=True)

        # Tabla resumen mensual
        with st.expander("📋 Detalle mes a mes", expanded=False):
            tbl_evol = evol.copy()
            tbl_evol["periodo"]     = tbl_evol["periodo"].dt.strftime("%b %Y")
            tbl_evol["unidades"]    = tbl_evol["unidades"].apply(lambda x: f"{x:,.0f}")
            tbl_evol["facturacion"] = tbl_evol["facturacion"].apply(fmt_peso)
            tbl_evol["clientes"]    = tbl_evol["clientes"].astype(int)
            tbl_evol.columns        = ["Período","Unidades","Facturación","Clientes únicos"]
            st.dataframe(tbl_evol, use_container_width=True, hide_index=True)

        # Ranking de clientes
        st.markdown("---")
        st.markdown("#### 👥 Clientes que compran este artículo")
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
            top10_cli = top_cli.head(10)

            cl1, cl2 = st.columns([2, 1])
            with cl1:
                fig_cli = px.bar(
                    top10_cli, x=col_ord, y="cliente", orientation="h",
                    title="Top 10 clientes",
                    labels={col_ord: crit, "cliente": ""},
                    color_discrete_sequence=["#0066cc"],
                    text=top10_cli[col_ord].apply(
                        lambda x: f"{x:,.0f}" if crit == "Unidades" else fmt_compacto(x)))
                fig_cli.update_traces(textposition="outside")
                fig_cli.update_layout(yaxis={"categoryorder":"total ascending"},
                                      height=300, margin=dict(t=40, r=80))
                st.plotly_chart(fig_cli, use_container_width=True)
            with cl2:
                tbl_cli = top_cli[["#","cliente","unidades","facturacion","ultima_compra"]].copy()
                tbl_cli["unidades"]      = tbl_cli["unidades"].apply(lambda x: f"{x:,.0f}")
                tbl_cli["facturacion"]   = tbl_cli["facturacion"].apply(fmt_peso)
                tbl_cli["ultima_compra"] = tbl_cli["ultima_compra"].dt.strftime("%d/%m/%Y")
                tbl_cli.columns = ["#","Cliente","Uds","Facturación","Última compra"]
                st.dataframe(tbl_cli, use_container_width=True, hide_index=True, height=300)

            st.download_button(
                "📥 Descargar ranking",
                top_cli.to_csv(index=False).encode("utf-8"),
                file_name=f"clientes_{r['cod_articulo']}.csv",
                mime="text/csv", key="cli_dl")


# ── App principal ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📦 Gestión de Artículos")
    st.markdown("**Distribuidora Cómodo**")
    st.markdown("---")
    pwd = st.text_input("Contraseña:", type="password")
    st.markdown("---")
    if st.button("🔄 Actualizar datos", use_container_width=True,
                 help="Recarga el archivo desde Google Drive"):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")
    st.caption("📁 Datos desde Google Drive")

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

st.title("📦 Gestión de Artículos — Cómodo")
st.caption(f"Datos al **{hoy.strftime('%d/%m/%Y')}**")
st.divider()

vista_articulos(df_ventas, df_art=df_art, df_stock=df_stock, df_precios=df_precios)
