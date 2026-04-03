# pi_streamlit_final_pdf.py
# Streamlit app: Planned vs Delivered vs Pending by Team with high-quality PDF export
#
# Install dependencies:
# pip install streamlit pandas plotly kaleido PyPDF2 pillow openpyxl python-dateutil
#
# Run:
# streamlit run pi_streamlit_final_pdf.py

import io
import math
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from dateutil import parser
from PIL import Image
from PyPDF2 import PdfMerger

st.set_page_config(page_title="PI Planned vs Delivered ", layout="wide")
st.title("Planned vs Delivered vs Pending - PI Data")

# -------------------------
# Helpers: data detection & aggregation
# -------------------------
def try_parse_date_series(s):
    try:
        return pd.to_datetime(s, errors='coerce')
    except Exception:
        return pd.Series([pd.NaT] * len(s))

def compute_aggregated_from_raw(df):
    """
    Detect Team, Planned, Delivered, Pending columns (common names).
    Compute missing Pending or Delivered if possible.
    Return aggregated DataFrame (Team, Planned, Delivered, Pending, DeliveredPct) and the normalized raw df.
    """
    cols = [c.strip().lower() for c in df.columns]
    def find(names):
        for n in names:
            if n in cols:
                return df.columns[cols.index(n)]
        return None

    team_col = find(['team', 'team name', 'team_name', 'pi', 'pi name'])
    planned_col = find(['planned', 'planned stories', 'planned_count', 'planned_count'])
    delivered_col = find(['delivered', 'done', 'completed', 'delivered_count'])
    pending_col = find(['pending', 'remaining', 'backlog'])
    date_col = find(['date', 'completed date', 'done date', 'pi start date', 'pi_end_date', 'pi end date'])

    # Try numeric branches
    if planned_col and delivered_col:
        df['_planned'] = pd.to_numeric(df[planned_col], errors='coerce').fillna(0).astype(int)
        df['_delivered'] = pd.to_numeric(df[delivered_col], errors='coerce').fillna(0).astype(int)
        df['_pending'] = (df['_planned'] - df['_delivered']).clip(lower=0).astype(int)
    elif planned_col and pending_col:
        df['_planned'] = pd.to_numeric(df[planned_col], errors='coerce').fillna(0).astype(int)
        df['_pending'] = pd.to_numeric(df[pending_col], errors='coerce').fillna(0).astype(int)
        df['_delivered'] = (df['_planned'] - df['_pending']).clip(lower=0).astype(int)
    else:
        # fallback: exact column names
        fallback = {c.lower(): c for c in df.columns}
        if 'planned' in fallback and 'delivered' in fallback:
            df['_planned'] = pd.to_numeric(df[fallback['planned']], errors='coerce').fillna(0).astype(int)
            df['_delivered'] = pd.to_numeric(df[fallback['delivered']], errors='coerce').fillna(0).astype(int)
            df['_pending'] = (df['_planned'] - df['_delivered']).clip(lower=0).astype(int)
        else:
            raise ValueError("Could not detect numeric Planned/Delivered columns. Provide aggregated CSV or map columns.")

    if not team_col:
        df['_team'] = df.index.astype(str)
    else:
        df['_team'] = df[team_col].astype(str).str.strip()

    if date_col:
        df['_date'] = try_parse_date_series(df[date_col])
    else:
        df['_date'] = pd.NaT

    agg = df.groupby('_team').agg(
        Planned = ('_planned', 'sum'),
        Delivered = ('_delivered', 'sum'),
        Pending = ('_pending', 'sum')
    ).reset_index().rename(columns={'_team': 'Team'})
    agg['DeliveredPct'] = (agg['Delivered'] / agg['Planned'].replace({0: None}) * 100).fillna(0).round(1)
    return agg, df

# -------------------------
# Helpers: figure builders
# -------------------------
def build_stacked_fig(agg, colors, title):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=agg['Team'], y=agg['Delivered'], name='Delivered', marker_color=colors['delivered'],
        text=agg['Delivered'], textposition='inside', hovertemplate='%{x}<br>Delivered: %{y}<extra></extra>'
    ))
    fig.add_trace(go.Bar(
        x=agg['Team'], y=agg['Pending'], name='Pending', marker_color=colors['pending'],
        text=agg['Pending'], textposition='inside', hovertemplate='%{x}<br>Pending: %{y}<extra></extra>'
    ))
    fig.update_layout(
        barmode='stack',
        title=title,
        xaxis_title='Team',
        yaxis_title='Stories',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        template='simple_white',
        margin=dict(t=80, b=40, l=40, r=20),
        height=520
    )
    # Delivered % badges above each bar
    for i, row in agg.iterrows():
        total = row['Planned']
        pct = f"{row['DeliveredPct']:.1f}%"
        fig.add_annotation(x=row['Team'], y=total + max(1, total * 0.03), text=pct, showarrow=False,
                           font=dict(size=11, color='black'), bgcolor='white', bordercolor='lightgray', borderwidth=1)
    return fig

def build_grouped_fig(agg, colors, title):
    x = agg['Team']
    fig = go.Figure()
    fig.add_trace(go.Bar(x=x, y=agg['Planned'], name='Planned', marker_color=colors['planned'], text=agg['Planned'], textposition='outside'))
    fig.add_trace(go.Bar(x=x, y=agg['Delivered'], name='Delivered', marker_color=colors['delivered'], text=agg['Delivered'], textposition='outside'))
    fig.add_trace(go.Bar(x=x, y=agg['Pending'], name='Pending', marker_color=colors['pending'], text=agg['Pending'], textposition='outside'))
    fig.add_trace(go.Scatter(x=x, y=agg['DeliveredPct'], name='Delivered %', yaxis='y2',
                             mode='lines+markers+text', line=dict(color='black', dash='dash'),
                             text=[f"{v:.1f}%" for v in agg['DeliveredPct']], textposition='top center'))
    fig.update_layout(
        title=title,
        xaxis_title='Team',
        yaxis=dict(title='Stories'),
        yaxis2=dict(title='Delivered %', overlaying='y', side='right', range=[0, 110]),
        barmode='group',
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        template='simple_white',
        margin=dict(t=80, b=40, l=40, r=60),
        height=520
    )
    return fig

# -------------------------
# Helpers: export utilities (vector PDF preferred, fallback to high-res PNG->PDF)
# -------------------------
def fig_to_pdf_bytes_vector(fig):
    """
    Export Plotly figure to vector PDF bytes using kaleido.
    Raises exception if not available.
    """
    return fig.to_image(format="pdf", width=1200, height=700, scale=1)

def fig_to_png_bytes_highres(fig, scale=3):
    """
    Export Plotly figure to high-resolution PNG bytes (fallback).
    """
    return fig.to_image(format="png", width=1200, height=700, scale=scale)

def png_bytes_to_pdf_bytes(png_bytes):
    """
    Convert PNG bytes to single-page PDF bytes using PIL.
    """
    img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    pdf_io = io.BytesIO()
    img.save(pdf_io, format="PDF", resolution=300)
    pdf_io.seek(0)
    return pdf_io.read()

def merge_pdf_bytes(list_of_pdf_bytes):
    """
    Merge multiple PDF byte blobs into a single PDF byte blob using PyPDF2.
    """
    merger = PdfMerger()
    for b in list_of_pdf_bytes:
        merger.append(io.BytesIO(b))
    out_io = io.BytesIO()
    merger.write(out_io)
    merger.close()
    out_io.seek(0)
    return out_io.read()

# -------------------------
# UI: upload / sample
# -------------------------
uploaded = st.file_uploader("Upload aggregated CSV or raw Excel (Team, Planned, Delivered, Pending)", type=["csv", "xlsx", "xls"])
use_sample = st.checkbox("Use sample data instead", value=False)

if use_sample and not uploaded:
    sample_df = pd.DataFrame({
        "PI Start Date": ["April-1-2026"] * 5,
        "PI End Date": ["May-31-2026"] * 5,
        "Team": ["A", "B", "C", "D", "E"],
        "Planned": [16, 4, 20, 30, 21],
        "Delivered": [10, 3, 10, 25, 21],
        "Pending": [6, 1, 10, 5, 0]
    })
    agg, raw = compute_aggregated_from_raw(sample_df)
elif uploaded:
    try:
        if uploaded.name.lower().endswith(('.xls', '.xlsx')):
            raw_df = pd.read_excel(uploaded, sheet_name=0, engine='openpyxl')
            agg, raw = compute_aggregated_from_raw(raw_df)
        else:
            raw_df = pd.read_csv(uploaded)
            agg, raw = compute_aggregated_from_raw(raw_df)
    except Exception as e:
        st.error(f"Failed to parse uploaded file: {e}")
        st.stop()
else:
    st.info("Upload a file or check 'Use sample data' to proceed.")
    st.stop()

# -------------------------
# Controls
# -------------------------
st.sidebar.header("Display & Export options")
order_by = st.sidebar.selectbox("Order teams by", options=["Team (A→Z)", "Delivered % (desc)"])
highlight_threshold = st.sidebar.slider("Highlight teams with Delivered % below", min_value=0, max_value=100, value=60)
export_mode = st.sidebar.selectbox("PDF export mode", options=["Vector PDF (preferred)", "High‑DPI PNG→PDF (fallback)"], index=0)
export_scale = st.sidebar.selectbox("PNG export scale (for fallback)", options=[2, 3, 4], index=1)
show_summary = st.sidebar.checkbox("Show summary panel", value=True)

# ordering
if order_by == "Delivered % (desc)":
    agg = agg.sort_values('DeliveredPct', ascending=False).reset_index(drop=True)
else:
    team_order = ["A", "B", "C", "D", "E"]
    present = [t for t in team_order if t in agg['Team'].tolist()]
    if set(present) == set(agg['Team'].tolist()):
        agg['Team'] = pd.Categorical(agg['Team'], categories=present, ordered=True)
        agg = agg.sort_values('Team').reset_index(drop=True)
    else:
        agg = agg.sort_values('Team').reset_index(drop=True)

colors = {'planned': '#d3d3d3', 'delivered': '#2ca02c', 'pending': '#d62728'}

# -------------------------
# Layout: summary + charts
# -------------------------
left, right = st.columns([1, 2])
if show_summary:
    with left:
        st.subheader("PI Summary")
        total_planned = int(agg['Planned'].sum())
        total_delivered = int(agg['Delivered'].sum())
        total_pending = int(agg['Pending'].sum())
        overall_pct = (total_delivered / total_planned * 100) if total_planned else 0.0
        st.metric("Total Planned", f"{total_planned}")
        st.metric("Total Delivered", f"{total_delivered}", delta=f"{overall_pct:.1f}% delivered")
        st.metric("Total Pending", f"{total_pending}")
        display = agg.copy()
        display['DeliveredPctStr'] = display['DeliveredPct'].map(lambda v: f"{v:.1f}%")
        st.table(display[['Team', 'Planned', 'Delivered', 'Pending', 'DeliveredPctStr']].set_index('Team'))

with right:
    st.subheader("Charts")
    fig_stack = build_stacked_fig(agg, colors, "Planned vs Delivered vs Pending by Team (PI Apr–May 2026)")
    # highlight low performers
    for i, row in agg.iterrows():
        if row['DeliveredPct'] < highlight_threshold:
            fig_stack.add_shape(type="rect",
                                x0=i - 0.4, x1=i + 0.4,
                                y0=0, y1=row['Planned'] + max(1, row['Planned'] * 0.05),
                                xref='x', yref='y',
                                line=dict(color='orange', width=2),
                                fillcolor='rgba(0,0,0,0)')
    st.plotly_chart(fig_stack, use_container_width=True)

    fig_group = build_grouped_fig(agg, colors, "Planned / Delivered / Pending with Delivered %")
    st.plotly_chart(fig_group, use_container_width=True)

# -------------------------
# Exports: PNGs and merged PDF (vector preferred)
# -------------------------
st.subheader("Export")
col1, col2, col3 = st.columns(3)

with col1:
    try:
        png1 = fig_stack.to_image(format="png", width=1200, height=700, scale=2)
        st.download_button("Download stacked chart PNG", data=png1, file_name="team_planned_delivered_pending.png", mime="image/png")
    except Exception:
        st.warning("PNG export requires 'kaleido'. Install it to enable PNG downloads.")

with col2:
    try:
        png2 = fig_group.to_image(format="png", width=1200, height=700, scale=2)
        st.download_button("Download grouped chart PNG", data=png2, file_name="team_grouped_planned_delivered_pending.png", mime="image/png")
    except Exception:
        st.warning("PNG export requires 'kaleido'. Install it to enable PNG downloads.")

with col3:
    # Build merged PDF using vector export if possible, else fallback to high-res PNG->PDF
    try:
        pdf_parts = []
        if export_mode == "Vector PDF (preferred)":
            try:
                pdf1 = fig_stack.to_image(format="pdf", width=1200, height=700, scale=1)
                pdf2 = fig_group.to_image(format="pdf", width=1200, height=700, scale=1)
                pdf_parts = [pdf1, pdf2]
            except Exception:
                # fallback to high-res PNG->PDF
                st.warning("Vector PDF export failed; falling back to high‑DPI PNG→PDF.")
                png1 = fig_stack.to_image(format="png", width=1200, height=700, scale=export_scale)
                png2 = fig_group.to_image(format="png", width=1200, height=700, scale=export_scale)
                pdf_parts = [png_bytes_to_pdf_bytes(png1), png_bytes_to_pdf_bytes(png2)]
        else:
            png1 = fig_stack.to_image(format="png", width=1200, height=700, scale=export_scale)
            png2 = fig_group.to_image(format="png", width=1200, height=700, scale=export_scale)
            pdf_parts = [png_bytes_to_pdf_bytes(png1), png_bytes_to_pdf_bytes(png2)]

        merged_pdf = merge_pdf_bytes(pdf_parts)
        st.download_button("Download combined PDF (high quality)", data=merged_pdf, file_name="pi_infographic_highquality.pdf", mime="application/pdf")
        st.success("PDF ready. Vector export preserves sharpness; fallback uses high‑DPI PNGs.")
    except Exception as e:
        st.error(f"PDF export failed: {e}. Ensure 'kaleido' and 'PyPDF2' are installed.")

# aggregated CSV download
st.subheader("Aggregated Data")
st.dataframe(agg.style.format({"DeliveredPct": "{:.1f}%"}))
csv_bytes = agg.to_csv(index=False).encode('utf-8')
st.download_button("Download aggregated CSV", data=csv_bytes, file_name="aggregated_pi.csv", mime="text/csv")

# -------------------------
# Final quick insight
# -------------------------
st.markdown("### Quick insight")
team_d_delivered = int(agg.loc[agg['Team'] == 'D', 'Delivered'].values[0]) if 'D' in agg['Team'].values else int(agg['Delivered'].max())
team_c_pending = int(agg.loc[agg['Team'] == 'C', 'Pending'].values[0]) if 'C' in agg['Team'].values else int(agg['Pending'].max())
overall_pct = (agg['Delivered'].sum() / agg['Planned'].sum() * 100) if agg['Planned'].sum() else 0.0
st.write(f"**Team D** delivered the largest absolute number vs plan ({team_d_delivered} delivered). "
         f"**Team C** has the highest pending ({team_c_pending}). "
         f"Overall delivered percentage across all teams is **{overall_pct:.1f}%**.")

st.success("Use the sidebar options to adjust ordering, highlight threshold, and export mode. Increase PNG scale for even higher resolution if needed.")