import csv
import io
import re
import unicodedata
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Relatório de Médicos por Setor", page_icon="logo.png", layout="wide")
st.markdown('''<style>
:root{--navy:#16324F;--teal:#0F766E;--ink:#172033}
.stApp{background:linear-gradient(180deg,#F8FAFC 0%,#FFF 38%);color:var(--ink)}
[data-testid="stSidebar"]{background:linear-gradient(180deg,#102A43,#16324F)}
[data-testid="stSidebar"] *{color:#F8FAFC!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section{background:#FFF!important;border:1px solid #CBD5E1!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"] section div{color:#16324F!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"] button{color:#16324F!important;background:#F8FAFC!important;border:1px solid #94A3B8!important}
[data-testid="stSidebar"] [data-testid="stFileUploader"] button *{color:#16324F!important}
[data-testid="stSidebar"] [data-baseweb="select"]>div{background:#FFF!important;border-color:#94A3B8!important}
[data-testid="stSidebar"] [data-baseweb="select"] *,[data-testid="stSidebar"] input{color:#16324F!important;-webkit-text-fill-color:#16324F!important}
.hero{padding:26px 30px;border-radius:20px;margin:4px 0 24px;color:#FFF;background:radial-gradient(circle at 90% 10%,#2DD4BF 0,transparent 30%),linear-gradient(120deg,#16324F,#0F766E);box-shadow:0 12px 28px #0f766e2e}
.hero h1{margin:0;font-size:2rem}.hero p{margin:8px 0 0;color:#D9FDF5}
div[data-testid="stMetric"]{background:#FFF;border:1px solid #E2E8F0;border-radius:14px;padding:12px;box-shadow:0 5px 16px #0f172a0d}
div[data-testid="stMetricValue"]{color:#16324F}.stButton button,.stDownloadButton button{border-radius:10px;font-weight:650}
</style>''', unsafe_allow_html=True)

EXTENSIONS=["csv","xls","xlsx","xlsm"]
ALIASES={
 "numero funcional":"NUMFUNC","numero do funcionario":"NUMFUNC","numero vinculo":"NUMVINC",
 "nome servidor":"SERVIDOR","descricao escala":"DESC. ESCALA","desc escala":"DESC. ESCALA",
 "setor":"SETOR","cargo":"CARGO","ocupacao":"OCUPACAO","cpf":"CPF",
 "carga horaria escalada":"CARGA HORARIA ESCALADA","carga horaria":"CARGA HORARIA","escala":"ESCALA",
}
REQUIRED=["NUMFUNC","NUMVINC","SETOR","CARGO"]
REPORT_COLUMNS=["ORDEM","ESCALA","DESC. ESCALA","NUMFUNC","NUMVINC","SERVIDOR","CPF","SETOR","CARGO","OCUPACAO","CARGA HORARIA ESCALADA","CARGA HORARIA"]


def text(v):
 if v is None:return ""
 try:
  if pd.isna(v):return ""
 except (TypeError,ValueError):pass
 s=str(v).strip(); return "" if s.casefold() in {"nan","none","nat"} else s

def key(v):
 s="".join(c for c in unicodedata.normalize("NFKD",text(v).casefold()) if not unicodedata.combining(c))
 return re.sub(r"[^a-z0-9]+"," ",s).strip()

def norm(v): return key(v).upper()

def read_raw(data,name):
 if name.casefold().endswith(".csv"):
  for enc in ("utf-8-sig","utf-8","cp1252","latin1"):
   try: raw=data.decode(enc);break
   except UnicodeDecodeError: continue
  try: sep=csv.Sniffer().sniff(raw[:10000],delimiters="\t;,|").delimiter
  except csv.Error: sep=";"
  return pd.read_csv(io.StringIO(raw),sep=sep,header=None,dtype=object,keep_default_na=False)
 return pd.read_excel(BytesIO(data),engine="xlrd" if data[:8]==b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" else "openpyxl",header=None,dtype=object)

def parse(raw):
 raw=raw.dropna(how="all").reset_index(drop=True)
 if raw.empty:return pd.DataFrame()
 idx=0
 for i in range(min(20,len(raw))):
  vals={key(x).replace(" ","").upper() for x in raw.iloc[i].tolist()}
  if "CARGO" in vals and "SETOR" in vals:idx=i;break
 body=raw.iloc[idx+1:].copy(); cols=[]; used=set()
 for value in raw.iloc[idx].tolist():
  label=text(value); canonical=ALIASES.get(key(label),label.upper())
  if canonical in used:
   n=2
   while f"{canonical}_{n}" in used:n+=1
   canonical=f"{canonical}_{n}"
  used.add(canonical);cols.append(canonical)
 body.columns=cols; return body.reset_index(drop=True)

def clean(df):
 df=df.copy()
 for c in df.columns: df[c]=df[c].map(text)
 for c in ["NUMFUNC","NUMVINC"]:
  if c not in df:df[c]=""
  df[c]=df[c].map(lambda v: re.sub(r"\D","",text(v)).lstrip("0") or "0" if re.sub(r"\D","",text(v)) else "")
 for c in REQUIRED:
  if c not in df:df[c]=""
 df["CARGO_NORM"]=df["CARGO"].map(norm)
 df["SETOR"]=df["SETOR"].map(text)
 return df[df["CARGO_NORM"].isin({"MEDICO","MEDICO RQE","MEDICO CLINICO"})].reset_index(drop=True)

def report(df):
 out=df.copy()
 for c in REPORT_COLUMNS:
  if c not in out:out[c]=""
 out=out[REPORT_COLUMNS].reset_index(drop=True);out["ORDEM"]=range(1,len(out)+1);return out

def summary_sector(df):
 if df.empty:return pd.DataFrame(columns=["SETOR","MEDICOS","VINCULOS","ESPECIALIDADES"])
 out=df.assign(_ESP=df["OCUPACAO"].replace("","Não informada"))
 return out.groupby("SETOR",dropna=False).agg(MEDICOS=("NUMFUNC","nunique"),VINCULOS=("NUMVINC","count"),ESPECIALIDADES=("_ESP",lambda x:"; ".join(sorted(set(x))))).reset_index().sort_values("SETOR")

def summary_specialty(df):
 if df.empty:return pd.DataFrame(columns=["SETOR","ESPECIALIDADE","MEDICOS"])
 out=df.assign(ESPECIALIDADE=df["OCUPACAO"].replace("","Não informada"))
 return out.groupby(["SETOR","ESPECIALIDADE"],dropna=False).agg(MEDICOS=("NUMFUNC","nunique")).reset_index().sort_values(["SETOR","ESPECIALIDADE"])

def style(ws,df):
 fill=PatternFill("solid",fgColor="16324F");font=Font(bold=True,color="FFFFFF");side=Side(style="thin",color="CBD5E1")
 for i,c in enumerate(df.columns,1):
  cell=ws.cell(1,i);cell.fill=fill;cell.font=font;cell.alignment=Alignment(horizontal="center",wrap_text=True);cell.border=Border(left=side,right=side,top=side,bottom=side)
  vals=[len(str(c))]+([int(df[c].astype(str).map(len).max())] if len(df) else []);ws.column_dimensions[get_column_letter(i)].width=min(max(max(vals)+2,12),50)
 for row in ws.iter_rows(min_row=2,max_row=ws.max_row,max_col=len(df.columns)):
  for cell in row:cell.border=Border(left=side,right=side,top=side,bottom=side);cell.alignment=Alignment(vertical="center")
 if len(df):ws.freeze_panes="A2";ws.auto_filter.ref=ws.dimensions;ws.print_title_rows="1:1";ws.print_area=f"A1:{get_column_letter(len(df.columns))}{ws.max_row}"
 ws.page_setup.orientation="landscape";ws.page_setup.paperSize=ws.PAPERSIZE_A4;ws.page_setup.fitToWidth=1;ws.page_setup.fitToHeight=0;ws.sheet_properties.pageSetUpPr.fitToPage=True

def excel(sheets):
 out=BytesIO()
 with pd.ExcelWriter(out,engine="openpyxl") as writer:
  for name,df in sheets.items():df.to_excel(writer,sheet_name=name[:31],index=False);style(writer.sheets[name[:31]],df)
 return out.getvalue()

def pdf(summary, detail, specialties):
 from reportlab.lib import colors
 from reportlab.lib.pagesizes import A4,landscape
 from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
 from reportlab.lib.units import mm
 from reportlab.platypus import SimpleDocTemplate,Paragraph,Spacer,Table,TableStyle
 from xml.sax.saxutils import escape
 out=BytesIO();doc=SimpleDocTemplate(out,pagesize=landscape(A4),leftMargin=12*mm,rightMargin=12*mm,topMargin=10*mm,bottomMargin=10*mm,title="Resumo de médicos")
 styles=getSampleStyleSheet(); title=ParagraphStyle("t",parent=styles["Title"],fontSize=18,textColor=colors.HexColor("#16324F")); cell=ParagraphStyle("c",parent=styles["Normal"],fontSize=7,leading=8);head=ParagraphStyle("h",parent=cell,textColor=colors.white,fontName="Helvetica-Bold")
 story=[Paragraph("Resumo estatístico — Médicos por setor",title),Paragraph(f"Gerado em {datetime.today().strftime('%d/%m/%Y %H:%M')} | Cargos considerados: MÉDICO, MEDICO - RQE e MEDICO CLINICO",styles["Normal"]),Spacer(1,6*mm)]
 data=[[Paragraph(x,head) for x in ["Setor","Médicos","Vínculos","Especialidades"]]]
 for _,r in summary.iterrows():data.append([Paragraph(escape(text(r.SETOR)),cell),text(r.MEDICOS),text(r.VINCULOS),Paragraph(escape(text(r.ESPECIALIDADES)),cell)])
 if len(data)==1:data.append([Paragraph("Nenhum médico encontrado.",cell),"","",""])
 t=Table(data,colWidths=[75*mm,25*mm,25*mm,130*mm],repeatRows=1);t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#0F766E")),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#CBD5E1")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F8FAFC")]),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
 story += [t,Spacer(1,6*mm),Paragraph("Especialidades por setor",styles["Heading2"])]
 data2=[[Paragraph(x,head) for x in ["Setor","Especialidade","Médicos"]]]
 for _,r in specialties.iterrows():data2.append([Paragraph(escape(text(r.SETOR)),cell),Paragraph(escape(text(r.ESPECIALIDADE)),cell),text(r.MEDICOS)])
 t2=Table(data2,colWidths=[100*mm,120*mm,35*mm],repeatRows=1);t2.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#16324F")),("GRID",(0,0),(-1,-1),.35,colors.HexColor("#CBD5E1")),("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F8FAFC")])]))
 story.append(t2);doc.build(story);return out.getvalue()

def main():
 st.markdown('<div class="hero"><div>RELATÓRIO DE CBO</div><h1>Médicos por setor</h1><p>Relação de médicos e especialidades cadastradas em cada setor.</p></div>',unsafe_allow_html=True)
 with st.sidebar:
  st.header("Arquivo da Intranet");upload=st.file_uploader("Planilha de escalas",type=EXTENSIONS);st.caption("São considerados somente os cargos MÉDICO, MEDICO - RQE e MEDICO CLINICO.")
 if not upload:st.info("Envie a planilha da Intranet para gerar o relatório.");st.stop()
 try:
  raw=parse(read_raw(upload.getvalue(),upload.name));missing=[c for c in REQUIRED if c not in raw.columns]
  if missing:st.error("Colunas obrigatórias ausentes: "+", ".join(missing));st.stop()
  doctors=clean(raw)
 except Exception as e:st.error(f"Não foi possível processar a planilha: {e}");st.stop()
 sectors=["(Todos)"]+sorted(doctors.SETOR.unique().tolist()) if len(doctors) else ["(Todos)"]
 with st.sidebar:sector=st.selectbox("Filtrar setor",sectors)
 filtered=doctors if sector=="(Todos)" else doctors[doctors.SETOR==sector]
 detail=report(filtered);sector_summary=summary_sector(doctors);specialty_summary=summary_specialty(doctors)
 k1,k2,k3=st.columns(3);k1.metric("Médicos",doctors.NUMFUNC.nunique());k2.metric("Vínculos",len(doctors));k3.metric("Setores",doctors.SETOR.nunique())
 tabs=st.tabs(["Relação de médicos","Resumo por setor","Especialidades por setor"])
 with tabs[0]:
  if detail.empty:st.warning("Nenhum médico encontrado com os três cargos definidos.")
  else:st.dataframe(detail,use_container_width=True,hide_index=True,height=520)
 with tabs[1]:st.dataframe(sector_summary,use_container_width=True,hide_index=True)
 with tabs[2]:st.dataframe(specialty_summary,use_container_width=True,hide_index=True)
 st.markdown("---");st.subheader("Exportar relatórios")
 xlsx=excel({"Medicos":detail,"Resumo setores":sector_summary,"Especialidades":specialty_summary});p=pdf(sector_summary,detail,specialty_summary)
 a,b=st.columns(2)
 with a:st.download_button("Baixar Excel",xlsx,f"relatorio_medicos_cbo_{datetime.today():%Y%m%d}.xlsx","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
 with b:st.download_button("Baixar resumo em PDF",p,f"resumo_medicos_cbo_{datetime.today():%Y%m%d}.pdf","application/pdf",use_container_width=True)
if __name__=="__main__":main()
