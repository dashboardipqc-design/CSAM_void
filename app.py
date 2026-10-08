import io
import re
import hashlib
from datetime import datetime
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title='CSAM Abnormality | Management Dashboard', page_icon='📊', layout='wide', initial_sidebar_state='expanded')
st.markdown('''<style>
.stApp{background:#090e17;color:#f1f5f9} [data-testid="stSidebar"]{background:#101a2b}
[data-testid="stMetric"]{background:#111b2d;border:1px solid #2a3c59;border-radius:14px;padding:17px}
[data-testid="stMetricLabel"]{color:#a9bbd3}h1,h2,h3{color:#f1f5f9}
[data-testid="stFileUploader"]{border-radius:12px} .block-container{padding-top:1.8rem}
</style>''', unsafe_allow_html=True)
COLORS=['#38bdf8','#a78bfa','#34d399','#fb923c','#f472b6','#facc15','#94a3b8']
DISPOSITIONS=['Accepted','Rework','Pending Verification','Rejected / Scrapped','Other / Unclassified']
DISP_COLORS={'Accepted':'#34d399','Rework':'#38bdf8','Pending Verification':'#facc15','Rejected / Scrapped':'#fb7185','Other / Unclassified':'#94a3b8'}
ALIASES={
 'Date':['date','inspection date','record date','reported date','event date'],
 'Product':['product','device','part','part number','model'],
 'Process':['process','area','department','section','station','operation'],
 'Lot':['lot','lot #','lot no','lot number','lot id','record id','id'],
 'Quantity':['qty','quantity','void qty','affected qty','affected quantity','units','no of units'],
 'Disposition':['disposition','final disposition','status','decision','final status'],
 'Remarks':['remarks','remark','comments','comment','notes','description','finding'],
}
def clean(v):
 if pd.isna(v): return ''
 return re.sub(r'\s+',' ',str(v)).strip()
def norm(v): return re.sub(r'[^a-z0-9]+',' ',clean(v).lower()).strip()
def detect(cols):
 mapping={}; names={norm(c):c for c in cols}
 for field,aliases in ALIASES.items():
  for alias in aliases:
   if norm(alias) in names: mapping[field]=names[norm(alias)];break
 return mapping
def disposition(remarks):
 s=norm(remarks)
 if not s:return 'Other / Unclassified'
 if re.search(r'\b(scrap|scrapped|reject|rejected|dispose|disposal)\b',s):return 'Rejected / Scrapped'
 if re.search(r'\b(rework|re worked|reworked|repair|repaired)\b',s):return 'Rework'
 if re.search(r'\b(pending|waiting|awaiting|hold|verification|verify|on going|ongoing|to confirm|for disposition|under investigation)\b',s):return 'Pending Verification'
 if re.search(r'\b(accept|accepted|approve|approved|use as is|proceed as normal|release|released)\b',s):return 'Accepted'
 return 'Other / Unclassified'
def dates(series):
 def one(v):
  if pd.isna(v) or clean(v)=='':return pd.NaT
  if isinstance(v,(int,float,np.integer,np.floating)) and 20000<float(v)<100000:return pd.Timestamp('1899-12-30')+pd.to_timedelta(float(v),unit='D')
  try:return pd.to_datetime(v,errors='coerce')
  except Exception:return pd.NaT
 return pd.to_datetime(series.map(one),errors='coerce')
@st.cache_data(show_spinner=False)
def workbook(file_bytes):
 xl=pd.ExcelFile(io.BytesIO(file_bytes));return xl.sheet_names
@st.cache_data(show_spinner=False)
def load_sheet(file_bytes,sheet):
 df=pd.read_excel(io.BytesIO(file_bytes),sheet_name=sheet,dtype=object)
 df=df.dropna(how='all').copy();df.columns=[clean(c) if clean(c) else f'Unnamed {i+1}' for i,c in enumerate(df.columns)]
 df=df.loc[:,~pd.Index(df.columns).duplicated()].copy()
 return df

def prepare(raw,sheet,override):
 mapping=detect(raw.columns)
 mapping.update({k:v for k,v in override.items() if v!='(not mapped)'})
 out=raw.copy()
 for f in ALIASES:
  out['__'+f]=raw[mapping[f]].map(clean) if f in mapping else ''
 if 'Process' not in mapping:out['__Process']=sheet if sheet.upper()!='ALL' else 'Unspecified'
 out['__Process']=out['__Process'].replace('', 'Unspecified').str.upper()
 out['__Date']=dates(raw[mapping['Date']]) if 'Date' in mapping else pd.Series(pd.NaT,index=out.index)
 out['__Month']=out['__Date'].dt.strftime('%Y-%m').fillna('Unknown')
 out['__Quantity']=pd.to_numeric(raw[mapping['Quantity']].map(lambda x: re.sub(r',','',clean(x))),errors='coerce') if 'Quantity' in mapping else pd.Series(np.nan,index=out.index)
 out['__DispositionSource']='Provisional (remarks)'
 if 'Disposition' in mapping:
  explicit=out['__Disposition'].map(norm)
  def direct(s):
   if not s:return None
   if re.search(r'\b(scrap|reject)\b',s):return 'Rejected / Scrapped'
   if re.search(r'\b(rework|repair)\b',s):return 'Rework'
   if re.search(r'\b(pending|hold|verify|verification|waiting)\b',s):return 'Pending Verification'
   if re.search(r'\b(accept|approve|release|pass|proceed)\b',s):return 'Accepted'
   return 'Other / Unclassified'
  mapped=explicit.map(direct)
  out['__Disposition']=mapped.where(mapped.notna(),out['__Remarks'].map(disposition))
  out['__DispositionSource']=np.where(mapped.notna(),'Explicit column','Provisional (remarks)')
 else:out['__Disposition']=out['__Remarks'].map(disposition)
 # Exact duplicates across source fields, retaining all records and original row numbers.
 source=raw.fillna('').astype(str).apply(lambda col:col.str.strip().str.lower())
 out['__Duplicate']=source.duplicated(keep='first')
 out['__ExcelRow']=raw.index+2
 return out,mapping

def fig_style(fig,height=350):
 fig.update_layout(template='plotly_dark',paper_bgcolor='#111b2d',plot_bgcolor='#111b2d',font_color='#e2e8f0',height=height,margin=dict(l=18,r=15,t=25,b=35),legend=dict(orientation='h',y=-.3,x=0),hovermode='x unified')
 fig.update_xaxes(gridcolor='#26374e', tickfont=dict(size=10), title_font=dict(size=11));fig.update_yaxes(gridcolor='#26374e', tickfont=dict(size=10), title_font=dict(size=11));fig.update_layout(font=dict(size=11),legend_title_font=dict(size=11),legend_font=dict(size=10))
 return fig

def chart_or_empty(frame,x,y,color,title,kind='bar'):
 st.subheader(title)
 if frame.empty:st.info('No matching records for this chart.');return
 if kind=='bar':fig=px.bar(frame,x=x,y=y,color=color,color_discrete_sequence=COLORS,barmode='stack')
 else:fig=px.line(frame,x=x,y=y,color=color,markers=True,color_discrete_sequence=COLORS)
 if x == '__Month':
  months=sorted(frame[x].dropna().unique())
  labels=[pd.to_datetime(m,format='%Y-%m').strftime("%b'%y") if m != 'Unknown' else 'Unknown' for m in months]
  fig.update_xaxes(type='category',categoryorder='array',categoryarray=months,tickmode='array',tickvals=months,ticktext=labels,tickangle=0,title_text='Month',tickfont=dict(size=9),automargin=True)
 if color == '__Process':fig.update_layout(legend_title_text='Process')
 st.plotly_chart(fig_style(fig),use_container_width=True)

st.title('CSAM Abnormality Monitoring')
st.caption('Occurrence • Affected quantity • Engineering disposition')
st.info('Accepted or reworked cases remain abnormality occurrences. Remarks-based disposition is provisional; missing rejection evidence does not establish confirmed zero rejection.')
with st.sidebar:
 st.header('Workbook & filters')
 upload=st.file_uploader('Upload Excel workbook',type=['xlsx','xls'],help='Uploaded workbook is processed by the Streamlit server for this session.')
if upload is None:
 st.markdown('### Upload an Excel workbook to begin')
 st.write('Supports multiple worksheets, flexible column matching, management KPIs, interactive charts, duplicate controls, and CSV export.')
 st.stop()
try:
 data=upload.getvalue();sheets=workbook(data)
 if not sheets:st.error('Workbook has no readable worksheets.');st.stop()
 default=next((i for i,s in enumerate(sheets) if s.strip().upper()=='ALL'),0)
 with st.sidebar:sheet=st.selectbox('Worksheet',sheets,index=default)
 raw=load_sheet(data,sheet)
 if raw.empty:st.warning('The selected worksheet has no records.');st.stop()
 auto=detect(raw.columns)
 with st.sidebar.expander('Column mapping (optional)'):
  overrides={}
  for field in ALIASES:
   options=['(not mapped)']+list(raw.columns)
   guess=auto.get(field,'(not mapped)')
   overrides[field]=st.selectbox(field,options,index=options.index(guess),key=f'map_{field}')
 df,mapping=prepare(raw,sheet,overrides)
except Exception as exc:
 st.error(f'Unable to read the selected Excel file: {exc}');st.stop()
with st.sidebar:
 if 'Date' not in mapping:st.warning('No date column mapped; month charts may show Unknown.')
 if 'Quantity' not in mapping:st.warning('No quantity column mapped; affected quantity cannot be calculated.')
 if 'Disposition' not in mapping:st.caption('Disposition inferred from remarks; all classifications are provisional.')
 processes=sorted(df['__Process'].unique())
 process=st.multiselect('Process / Area',processes,default=processes)
 disp=st.multiselect('Disposition',DISPOSITIONS,default=DISPOSITIONS)
 valid=df['__Date'].dropna()
 if not valid.empty:
  min_date,max_date=valid.min().date(),valid.max().date()
  selected=st.date_input('Date range',(min_date,max_date),min_value=min_date,max_value=max_date)
  include_undated=st.checkbox('Include records without valid dates',value=True)
 else:selected=None;include_undated=True
 duplicates=st.radio('Duplicates',['Include and flag','Exclude duplicates','Duplicates only'],index=0)
 if st.button('Reset filters',use_container_width=True):
  for k in list(st.session_state):
   if k.startswith('map_'):continue
   if k not in ('_uploaded_file',):del st.session_state[k]
  st.rerun()
mask=df['__Process'].isin(process)&df['__Disposition'].isin(disp)
if selected is not None and isinstance(selected,(tuple,list)) and len(selected)==2:
 start,end=selected
 mask &= df['__Date'].between(pd.Timestamp(start),pd.Timestamp(end)+pd.Timedelta(days=1)-pd.Timedelta(microseconds=1)) | (df['__Date'].isna() if include_undated else False)
if duplicates=='Exclude duplicates':mask &= ~df['__Duplicate']
elif duplicates=='Duplicates only':mask &= df['__Duplicate']
f=df.loc[mask].copy()
st.caption(f'Workbook: {upload.name}  •  Worksheet: {sheet}  •  {len(f):,} displayed of {len(df):,} source records  •  {int(df["__Duplicate"].sum()):,} potential exact duplicate rows')
quantity=f['__Quantity'].sum(min_count=1)
metrics=[('Total Cases',len(f)),('Affected Quantity',f'{quantity:,.0f}' if pd.notna(quantity) else 'N/A'),('Accepted Cases',int((f['__Disposition']=='Accepted').sum())),('Rework Cases',int((f['__Disposition']=='Rework').sum())),('Pending Cases',int((f['__Disposition']=='Pending Verification').sum())),('Rejected Cases',int((f['__Disposition']=='Rejected / Scrapped').sum()))]
for col,(label,value) in zip(st.columns(6),metrics):col.metric(label,value)
if f.empty:st.warning('No records match the selected filters. Adjust filters to display charts and records.')
else:
 left,right=st.columns(2)
 with left:
  monthly=f.groupby(['__Month','__Process'],dropna=False).size().reset_index(name='Cases').sort_values('__Month')
  chart_or_empty(monthly,'__Month','Cases','__Process','Monthly abnormality cases by process')
 with right:
  if f['__Quantity'].notna().any():
   q=f.groupby(['__Month','__Process'],dropna=False)['__Quantity'].sum().reset_index().sort_values('__Month')
   chart_or_empty(q,'__Month','__Quantity','__Process','Monthly affected quantity')
  else:st.subheader('Monthly affected quantity');st.info('No usable quantity values.')
 left,right=st.columns(2)
 with left:
  st.subheader('Disposition by case count')
  d=f['__Disposition'].value_counts().rename_axis('Disposition').reset_index(name='Cases')
  fig=px.pie(d,names='Disposition',values='Cases',hole=.63,color='Disposition',color_discrete_map=DISP_COLORS)
  st.plotly_chart(fig_style(fig),use_container_width=True)
 with right:
  p=f['__Process'].value_counts().rename_axis('Process').reset_index(name='Cases')
  chart_or_empty(p,'Process','Cases',None,'Process contribution')
 st.subheader('Disposition: cases vs affected quantity')
 summary=f.groupby('__Disposition',dropna=False).agg(Cases=('__Process','size'),Affected_Quantity=('__Quantity','sum')).reset_index().rename(columns={'__Disposition':'Disposition'})
 st.dataframe(summary.rename(columns={'Affected_Quantity':'Affected Quantity'}),hide_index=True,use_container_width=True)
 st.subheader('Management observations')
 top=f['__Process'].value_counts();pending=int((f['__Disposition']=='Pending Verification').sum());rework=int((f['__Disposition']=='Rework').sum());accepted=int((f['__Disposition']=='Accepted').sum())
 st.write(f'**Abnormality frequency:** {len(f):,} recorded cases within the selected scope. **Highest-contributing process:** {top.index[0]} ({top.iloc[0]:,} cases, {top.iloc[0]/len(f):.1%} of cases).')
 st.write(f'**Disposition mix:** {rework:,} rework cases ({rework/len(f):.1%}); {accepted:,} accepted cases ({accepted/len(f):.1%}); {pending:,} pending-verification cases ({pending/len(f):.1%}).')
 month_counts=f[f['__Month']!='Unknown'].groupby('__Month').size().sort_index()
 if len(month_counts)>1:st.write(f'**Monthly trend:** {month_counts.index[-1]} recorded {month_counts.iloc[-1]:,} cases versus {month_counts.iloc[-2]:,} in {month_counts.index[-2]}; partial-month coverage may affect comparison.')
 if (f['__DispositionSource']=='Provisional (remarks)').any():st.caption('Disposition summaries include provisional classifications inferred from remarks. Validate against engineering disposition records before final management conclusions.')
 if int((f['__Disposition']=='Rejected / Scrapped').sum())==0:st.caption('No rejected/scrapped classifications found in the filtered records; this is not confirmation that no material was rejected.')
st.subheader('Detailed abnormality register')
search=st.text_input('Search all record fields',placeholder='Search lot, product, remarks, process…')
register=f.copy()
if_search_cols=[c for c in raw.columns if c in register.columns]
if search:
 terms=register[if_search_cols].fillna('').astype(str).agg(' '.join,axis=1)
 register=register[terms.str.contains(re.escape(search),case=False,na=False)]
show_cols=list(raw.columns)+['__ExcelRow','__Process','__Date','__Quantity','__Disposition','__DispositionSource','__Duplicate']
view=register[show_cols].rename(columns={'__ExcelRow':'Excel Row','__Process':'Normalized Process','__Date':'Parsed Date','__Quantity':'Parsed Affected Quantity','__Disposition':'Classified Disposition','__DispositionSource':'Classification Source','__Duplicate':'Potential Duplicate'})
st.dataframe(view,use_container_width=True,hide_index=True,height=450)
st.download_button('⬇ Export filtered records (CSV)',data=view.to_csv(index=False).encode('utf-8-sig'),file_name=f'CSAM_{sheet}_filtered.csv',mime='text/csv',use_container_width=False)
st.caption('Duplicate detection flags exact repeated source rows (first occurrence retained as unique). The uploaded workbook is not altered. Streamlit server processes uploaded data; use a trusted deployment for confidential files.')
