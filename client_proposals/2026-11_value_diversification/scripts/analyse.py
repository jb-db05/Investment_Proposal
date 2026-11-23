import openpyxl, json, sys
wb=openpyxl.load_workbook(sys.argv[1],data_only=True); ws=wb.active
rows=list(ws.iter_rows(values_only=True))
SECTIONS={'Cash','Short-term instruments','Bonds','Equities','Structured products','Hedge Funds','Private Assets','Gold and other commodities','Other investments'}
COMMO_OVERRIDE=['CURRENT ACCOUNT IN XAG','CURRENT ACCOUNT IN XPT','BRGF WORLD MINING','VANECK GOLD MINERS','M&R EARTHS','DCI-XAGUSD']
# manual issuer-level region / sector / theme tags (desk judgement)
META={
 'SYZ AM(CH) 1 LIQUID':('US','Money market',''),
 'GOLDMAN FLR':('US','Financials',''),'GOLDMAN SACHS JAN33':('US','Financials',''),
 'VONOVIA':('Eurozone','Real estate',''),'TEVA':('EM','Health care',''),
 'PHILIP MORRIS':('US','Consumer staples',''),'META PLATFORMS':('US','Communication services','AI'),
 'AMAZON':('US','Consumer discretionary','AI'),'PKN ORLEN':('EM','Energy','Energy'),
 'BNY MELLON S/D':('Global','Diversified credit',''),'STORM II':('Global','Diversified credit',''),
 'SYZ EM BONDS':('EM','Diversified credit',''),
 'AIRBUS':('Eurozone','Industrials','Defence'),'ALLIANZ':('Eurozone','Financials',''),
 'BAE SYSTEMS':('UK','Industrials','Defence'),'BERKSHIRE':('US','Financials',''),
 'DASSAULT AVIATION':('Eurozone','Industrials','Defence'),'DECALIA AEGIS':('Global','Industrials','Defence'),
 'FTGF CB INF VAL':('Global','Utilities',''),'G SACHS US SMALL':('US','Diversified',''),
 'GOLDMAN EU HIGH DIV':('Eurozone','Diversified',''),'HANNOVER':('Eurozone','Financials',''),
 'DIGITAL SECURITY':('Global','Information technology','AI'),'MSCI EUROP MC':('Eurozone','Diversified',''),
 'LEONARDO (RAGGR)':('Eurozone','Industrials','Defence'),'LLOYDS':('UK','Financials',''),
 'POLAR BIOTECH':('Global','Health care',''),'SAAB':('Eurozone','Industrials','Defence'),
 'CHINA NEXT WINNERS':('EM','Information technology','AI'),'SWISSCOM':('Switzerland','Communication services',''),
 'UBS(CH) EQ-SWISS INC':('Switzerland','Diversified',''),'VANECK VECTORS DEF':('Global','Industrials','Defence'),
 'ZURICH INSURANCE':('Switzerland','Financials',''),
 'NESN':('Switzerland','Consumer staples',''),'ABBN-GALD':('Switzerland','Industrials',''),
 'NKY-SPX-SX5E':('Global','Diversified',''),'GOOGL-MSFT-AMZN':('US','Information technology','AI'),
 'LEONAI1':('Eurozone','Industrials','Defence'),'CLN MSFT':('US','Information technology','AI'),
 'MILLENNIUM':('Global','Hedge fund',''),
 'ACE BUYOUT':('Global','Private equity',''),'KONA':('Global','Private equity',''),'ECHO FUND':('Eurozone','Private equity',''),
 'LEFT LANE':('US','Private equity',''),'THEMIS':('Global','Private equity',''),
 'BITCOIN':('Global','Digital assets',''),
}
def meta(desc,sec,ccy):
    for k,v in META.items():
        if k in desc: return v
    return None
cur=None; out=[]
hdr=[i for i,r in enumerate(rows) if r[1]=="Last operation date"][0]
for r in rows[hdr+1:]:
    a=r[0]
    if a in SECTIONS: cur=a; continue
    if r[5] is None: continue
    w=r[21] or 0; val=r[7] or 0
    if w<=0: continue
    desc=r[5]; ccy=r[12]
    ac={'Cash':'Cash & money market','Short-term instruments':'Cash & money market','Bonds':'Fixed income','Equities':'Equities',
        'Structured products':'Structured products','Hedge Funds':'Hedge funds','Private Assets':'Private assets',
        'Gold and other commodities':'Commodities','Other investments':'Digital assets'}[cur]
    if any(k in desc for k in COMMO_OVERRIDE): ac='Commodities'
    # commodity sub-type
    sub=''
    if ac=='Commodities':
        if 'GOLD KG' in desc or 'KRUGERRAND' in desc: sub='Gold (physical)'
        elif 'GOLD MINERS' in desc: sub='Gold miners'
        else: sub='Other metals & commodities'
    if ac=='Commodities':
        region,sector,theme='Commodities (no region)','Commodities',''
    elif ac=='Cash & money market' and cur=='Cash':
        region={'CHF':'Switzerland','EUR':'Eurozone','GBP':'UK','USD':'US'}.get(ccy,'Other'); sector='Cash'; theme=''
    elif cur=='Short-term instruments' and 'Fid.' in desc:
        region={'EUR':'Eurozone','GBP':'UK'}[ccy]; sector='Money market'; theme=''
    else:
        m=meta(desc,cur,ccy)
        if m is None: print('NO META',desc); m=('Global','Other','')
        region,sector,theme=m
    perf_eur=r[19]; perf_qc=r[18]
    out.append(dict(section=cur,asset_class=ac,sub=sub,isin=r[2],desc=desc,ccy=ccy,value_eur=val,weight=w,
        perf_eur=perf_eur,perf_qc=perf_qc,region=region,sector=sector,theme=theme,last_op=str(r[1])[:10] if r[1] else None,
        is_account=desc.startswith(('CURRENT ACCOUNT','INCOME TO BE')) ))
tot=sum(o['value_eur'] for o in out)
print('positions',len(out),'total',round(tot),'wsum',round(sum(o['weight'] for o in out),3))
from collections import defaultdict
def agg(key,filt=lambda o:True,norm=False):
    d=defaultdict(float)
    for o in out:
        if filt(o): d[o[key]]+=o['weight']
    s=sum(d.values())
    if norm: d={k:v/s*100 for k,v in d.items()}
    return dict(sorted(d.items(),key=lambda x:-x[1]))
res=dict(total_eur=tot,n=len(out),holdings=out,
  asset_class=agg('asset_class'),
  commodity_sub=agg('sub',lambda o:o['asset_class']=='Commodities'),
  region=agg('region'),
  ccy=agg('ccy'),
  eq_sector=agg('sector',lambda o:o['asset_class']=='Equities',norm=True),
  eq_region=agg('region',lambda o:o['asset_class']=='Equities',norm=True),
  theme=agg('theme',lambda o:o['theme']),
)
for k in ['asset_class','commodity_sub','region','ccy','eq_sector','eq_region','theme']:
    print(k,{a:round(b,2) for a,b in res[k].items()})
perf=[o for o in out if not o['is_account'] or o['ccy'] in ('XAG','XPT')]
perf=[o for o in perf if o['section']!='Short-term instruments' and o['perf_eur'] is not None and o['value_eur']>1000]
perf.sort(key=lambda o:-o['perf_eur'])
print('\nTOP'); [print(round(o['perf_eur'],1),round(o['weight'],2),o['desc'],o['last_op']) for o in perf[:10]]
print('BOTTOM'); [print(round(o['perf_eur'],1),round(o['weight'],2),o['desc'],o['last_op']) for o in perf[-10:]]
res['perf_sorted']=[(o['desc'],o['asset_class'],o['weight'],o['perf_eur'],o['perf_qc'],o['last_op']) for o in perf]
json.dump(res,open('analysis.json','w'),indent=1,default=str)
