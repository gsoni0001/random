#!/usr/bin/env python3
"""
score_buyers.py — derive buyer-targeting scores from the classification/graph engine
+ optional signal feeds, and emit rows for the Master Targeting Table.

Inputs (CSV, swap in real data):
  package_profiles.csv   package tag profiles (union of asset tags) + offering facts
  buyer_portfolios.csv   tagged buyer portfolios + signal-feed columns
  tag_dictionary_updated.csv  (optional) tag specificity for IDF weighting

Derives per (package x target buyer): Strategic Fit + semantic fit (tag engine),
Citation Linkage (graph), FTO proxy, Litigation / Appetite / Standards / Financial
(feeds), Urgency (blend); runs the A-G sourcing detectors -> # signals + channel +
trigger + evidence; defaults the archetype approach columns. Marks every factor's
provenance tier (T1-T4 / proxy / blend).

Outputs: buyer_scores.csv (Master-table schema), scoring_provenance.csv, scoring_report.txt
"""
import csv, json, argparse, collections
W=[2,1,1,2,1,1.5,1,1.5]
import tagsim

def band(v, t5,t4,t3,t2):
    return 5 if v>=t5 else 4 if v>=t4 else 3 if v>=t3 else 2 if v>=t2 else (1 if v>0 else 0)
def tset(s):
    out=set()
    for t in [x.strip() for x in (s or '').split(';') if x.strip()]: out.update(tagsim._expand(t))
    return out
def dom_only(S): return {t for t in S if t.startswith('DOM.')}

ARCH={
 'OC-OFF':dict(deal='Outright sale or exclusive licence',ttc='Medium (3–5 mo)',
   hook='Product-mapping + EoU chart on their SKUs; quantify royalty base; survived-challenge de-risks',
   out='Direct — VP Patent Strategy',con='VP Patent Strategy',bundle='B26 EoU-Backed · B12 Detectability · B22 Anchor-Halo',nx='Send EoU teaser + NDA'),
 'OC-DEF':dict(deal='Outright sale (defensive shield)',ttc='Long (6–9 mo)',
   hook='Counter-assertion / FTO; clean title; deliver as a quiet sale notice (not a threat)',
   out='Direct — Head of IP / Litigation',con='Head of IP',bundle='B19 Defensive · B28 Clean-Title · B16 Foundational',nx='Quiet outreach via counsel'),
 'OC-EXP':dict(deal='Outright or field-of-use licence',ttc='Medium–Long (5–8 mo)',
   hook='Market-entry de-risking + convergence; cross-industry reads',
   out='Direct — Head of Domain IP',con='Head of Battery IP',bundle='B17 Cross-Industry · B30 Adjacent · B3 Product-Architecture',nx='Confirm build-vs-buy posture'),
 'IP-FUND':dict(deal='Outright sale',ttc='Medium (3–5 mo)',
   hook='Citation strength + adjacent re-read + resale optionality; portfolio IRR framing',
   out='Direct — Head of Acquisitions',con='Head of Acquisitions',bundle='B29 High-Citation · B30 Adjacent · B24 Strong-Core+Tail',nx='Send portfolio metrics pack'),
 'STD-LIC':dict(deal='Non-exclusive / FRAND licence',ttc='Long (9+ mo)',
   hook='SEP essentiality + FRAND leverage proportional to standards exposure',
   out='Standards & IP licensing arm',con='Head of Standards & IP',bundle='B2 SEP · B9 Interoperability · B23 Picket-Fence',nx='Assess SEP essentiality claim'),
 'DEF-AGG':dict(deal='Outright (volume lot, fixed price)',ttc='Short (1–3 mo)',
   hook='Volume / price + clean title; fast low-friction close; catch-and-release fit',
   out='AST IP3 / direct aggregator submission',con='Acquisitions Lead',bundle='B31 Salvage/Volume · B19 Defensive',nx='Submit to IP3 window'),
}
HEADWIND={0:'Low product exposure — weak FTO hook',1:'No litigation pressure — defensive only',2:'No citation linkage',
 3:'Weak domain / portfolio fit',4:'Limited standards exposure',5:'Price-sensitive / limited budget',
 6:'Low acquisition appetite',7:'No urgency — long timeline'}

def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--packages',default='package_profiles.csv')
    ap.add_argument('--buyers',default='buyer_portfolios.csv')
    ap.add_argument('--dict',default='tag_dictionary_updated.csv')
    a=ap.parse_args()
    spec=tagsim.load_specificity(a.dict)
    PK={r['package']:r for r in csv.DictReader(open(a.packages,newline=''))}
    for p in PK.values(): p['_set']=tset(p['tags'])
    buyers=list(csv.DictReader(open(a.buyers,newline='')))

    rows=[]
    for b in buyers:
        bset=tset(b['portfolio_tags']); pset=tset(b['product_tags'])
        prior=int(b['prior_buys']); suits=int(b['active_suits']); std=int(b['std_contribs'])
        launch=int(b['launch']); oa=int(b['oa_cite']); arch=b['archetype']
        cites=dict((x.split(':')[0], int(x.split(':')[1])) for x in (b['cites'] or '').split(';') if ':' in x)
        targets=[t for t in b['target_packages'].split(';') if t]
        for pkg in targets:
            P=PK[pkg]; psep=any(k in (P['sep'] or '') for k in ('SEP','3GPP')) or 'standard' in (P['sep'] or '').lower()
            j=tagsim.weighted_jaccard(bset,P['_set'],spec); sem=min(100,round(j*130))
            c=cites.get(pkg,0)
            # FTO proxy: product-domain overlap with package domain
            pdom=dom_only(pset); kdom=dom_only(P['_set'])
            fto_frac=(len(pdom & kdom)/len(kdom)*100) if (pdom and kdom) else 0
            # factors 0-5
            f_fit=band(sem,60,40,25,12)
            f_cite=band(c,3,2,1,1)
            f_fto=band(fto_frac,60,40,25,10)
            f_lit=band(suits,3,2,1,1)
            f_std=band(std,5,3,2,1) if psep else (2 if std>=3 else 1)
            f_fin=band(int(b['revenue']),10000,5000,2000,1000)
            if arch in ('IP-FUND','DEF-AGG','LIT-FIN'): f_fin=max(f_fin,4)
            f_appe=max(band(prior,10,5,2,1), 4 if arch in ('IP-FUND','DEF-AGG') else 2)
            f_urg=4 if launch else 2
            if suits>=1: f_urg=max(f_urg,5 if suits>=3 else 4)
            f=[f_fto,f_lit,f_cite,f_fit,f_std,f_fin,f_appe,f_urg]

            # A-G detectors
            fired=[]
            if prior>0: fired.append(('A','A · In-market','Assignment recordation mining',
                f'Recorded {prior} prior acquisitions in-class','USPTO Assignment DB'))
            if oa: fired.append(('B','B · Legal-pressure','Office-action citation (102 art)',
                'Anchor cited as 102 art against their pending application','Global Dossier'))
            elif suits>0: fired.append(('B','B · Legal-pressure','Litigation defendant',
                f'Named in {suits} active suit(s) in domain','Docket Navigator'))
            if sem>=40 or c>0 or (psep and std>=3):
                if sem>=40: ch,tg='Semantic portfolio overlap',f'IDF tag-overlap {sem}/100 with package'
                elif c>0: ch,tg='Forward-citation',f'Cites the anchor family x{c}'
                else: ch,tg='Standards-contribution gap',f'Heavy SSO contributor ({std}), thin SEP holdings'
                if c>0 and sem>=40: tg+=f'; cites anchor x{c}'
                fired.append(('C','C · Tech-trajectory',ch,tg,'Tag engine + USPTO'))
            if launch: fired.append(('D','D · Commercial-motion','Product launch / hiring',
                'New product line / hiring signal in domain','Press + LinkedIn'))
            if arch in ('IP-FUND','DEF-AGG','LIT-FIN'): fired.append(('E','E · Financial/structural','Acquisition mandate',
                'Standing fund/aggregator acquisition mandate','Fund/aggregator intake'))
            if len(targets)>1: fired.append(('F','F · Relationship','Cross-package adjacency',
                'Relevant to multiple packages (adjacency)','Cross-package match'))
            nsig=len(fired)
            order={'B':0,'C':1,'A':2,'D':3,'E':4,'F':5}
            prim=sorted(fired,key=lambda x:order[x[0]])[0] if fired else ('C','C · Tech-trajectory','Semantic overlap',f'Tag-overlap {sem}/100','Tag engine')

            comp=round(sum(W[i]*f[i] for i in range(8))/sum(W)*20)
            ad=ARCH.get(arch,ARCH['OC-DEF'])
            mn=min(range(8),key=lambda i:f[i])
            sev='H' if f[mn]<=1 else ('M' if f[mn]==2 else 'L')
            prod=(pset and sorted(pset)[0].replace('DOM.','').replace('.',' › ')) or 'n/a (financial buyer)'
            rows.append(dict(buyer=b['buyer'],pkg=pkg,arch=arch,sector=b['sector'],hq=b['hq'],
                rev=int(b['revenue']),ip=int(b['ip_count']),prior='Yes' if prior>0 else 'No',
                chan=prim[2],cat=prim[1],nsig=nsig,trig=prim[3],evid=prim[4],rec='Q1-2026',
                f0=f[0],f1=f[1],f2=f[2],f3=f[3],f4=f[4],f5=f[5],f6=f[6],f7=f[7],
                hw=HEADWIND[mn],sev=sev,da='M',dj='M' if suits>0 else 'L',
                bundle=ad['bundle'],sem=sem,prod=prod,deal=ad['deal'],ttc=ad['ttc'],hook=ad['hook'],
                out=ad['out'],con=ad['con'],stage='Not started',owner='GP' if ((comp>=75 and nsig>=2) or (comp>=65 and nsig>=3)) else 'Assoc',
                nx=ad['nx'],nxd='',note=f'Engine-scored · primary {prim[1]} · fit {sem}/100 · {nsig} signal(s)',_c=comp))

    rows.sort(key=lambda r:-r['_c'])
    cols=['buyer','pkg','arch','sector','hq','rev','ip','prior','chan','cat','nsig','trig','evid','rec',
          'f0','f1','f2','f3','f4','f5','f6','f7','hw','sev','da','dj','bundle','sem','prod','deal','ttc',
          'hook','out','con','stage','owner','nx','nxd','note']
    with open('buyer_scores.csv','w',newline='') as fo:
        w=csv.DictWriter(fo,fieldnames=cols,extrasaction='ignore'); w.writeheader(); w.writerows(rows)

    prov=[('Strategic Fit','T2','Tag engine — IDF-weighted tag overlap (tagsim)','derived'),
          ('Semantic fit (col)','T2','Tag engine — IDF overlap 0–100','derived'),
          ('Citation Linkage','T2','Citation graph (tag_graph) — buyer→asset citations','derived'),
          ('FTO Exposure','T2 → T3','Product/domain tag overlap = proxy; EoU claim-read to confirm','proxy'),
          ('Acquisition Appetite','T1','Assignment-recordation feed — prior buys','derived (feed)'),
          ('Litigation Pressure','T1','Litigation feed — active suits in domain','derived (feed)'),
          ('Standards / SEP','T1/T2','Package STD tags × standards-contribution feed','derived'),
          ('Financial Capacity','T4','Revenue / capital band (funds on AUM proxy)','derived (feed)'),
          ('Urgency','T1/T2','Blend: litigation + product-launch + remaining-term','blend'),
          ('# Signals / channel','T1/T2','A–G sourcing detectors over the feeds','derived')]
    with open('scoring_provenance.csv','w',newline='') as fo:
        w=csv.writer(fo); w.writerow(['Field','Tier','Source / method','Mode']); w.writerows(prov)

    catc=collections.Counter(r['cat'][:1] for r in rows)
    with open('scoring_report.txt','w') as fo:
        fo.write(f"Scored {len(rows)} (package × buyer) pairs from {len(buyers)} buyers, {len(PK)} packages.\n")
        fo.write(f"Composite range {min(r['_c'] for r in rows)}–{max(r['_c'] for r in rows)}.\n")
        fo.write(f"Multi-signal (≥2 channels): {sum(1 for r in rows if r['nsig']>=2)} / {len(rows)}.\n")
        fo.write(f"Primary signal category mix: {dict(catc)}\n")
        fo.write("Top 5 by composite:\n")
        for r in rows[:5]: fo.write(f"  {r['_c']:>3}  {r['buyer']:<24} {r['pkg']}  ({r['cat']}, {r['nsig']} sig)\n")
    print(f"OK: {len(rows)} rows -> buyer_scores.csv, scoring_provenance.csv, scoring_report.txt")

if __name__=='__main__': main()
