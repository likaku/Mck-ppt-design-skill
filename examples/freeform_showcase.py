#!/usr/bin/env python3
# Copyright 2024-2026 Kaku Li (https://github.com/likaku)
# Licensed under the Apache License, Version 2.0 — see LICENSE and NOTICE.
"""Free-composition example: the 6 slides in the README, drawn with python-pptx
primitives and native charts instead of the fixed layout methods.

Kept from the skill: full_cleanup() (no connectors, shadows or p:style), an East Asian
font on every run, 14pt body floor, source line and page number on every slide.

    python examples/freeform_showcase.py out.pptx
    python references/scripts/gate_check.py out.pptx .
"""
# -*- coding: utf-8 -*-
"""Free-composition build: same story as build.py, drawn with primitives + native chart.
Kept from the skill: full_cleanup (no p:style/shadow), East Asian font on every run, source line, page number, gate check."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
from pptx import Presentation
from pptx.util import Inches as I, Pt, Emu
from pptx.dml.color import RGBColor as C
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.oxml.ns import qn
from lxml import etree
from mck_ppt.core import full_cleanup

NAVY=C(0x05,0x1C,0x2C); INK=C(0x1F,0x29,0x37); GREY=C(0x6B,0x72,0x80); LIGHT=C(0xF3,0xF4,0xF6)
MID=C(0xD1,0xD5,0xDB); WHITE=C(0xFF,0xFF,0xFF); ACC=C(0x00,0x52,0xD9); TINT=C(0xE6,0xEE,0xFB); SOFT=C(0x9C,0xA3,0xAF)
CN='KaiTi'; LAT='Arial'; NUM='Consolas'
prs=Presentation(); prs.slide_width=I(13.333); prs.slide_height=I(7.5)
BLANK=prs.slide_layouts[6]
TOTAL=6

def rect(s,x,y,w,h,fill=None,line=None,lw=0.75):
    sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,I(x),I(y),I(w),I(h))
    if fill is None: sh.fill.background()
    else: sh.fill.solid(); sh.fill.fore_color.rgb=fill
    if line is None: sh.line.fill.background()
    else: sh.line.color.rgb=line; sh.line.width=Pt(lw)
    return sh
def shape(s,kind,x,y,w,h,fill):
    sh=s.shapes.add_shape(kind,I(x),I(y),I(w),I(h)); sh.fill.solid(); sh.fill.fore_color.rgb=fill; sh.line.fill.background(); return sh
def _run(p,text,size,color,bold,font):
    r=p.add_run(); r.text=text; f=r.font; f.size=Pt(size); f.bold=bold; f.color.rgb=color
    lat=font or LAT; f.name=lat
    rPr=r._r.get_or_add_rPr()
    for tag in ('a:ea','a:cs'):
        for e in rPr.findall(qn(tag)): rPr.remove(e)
        e=etree.SubElement(rPr,qn(tag)); e.set('typeface',CN)
    return r
def tx(s,x,y,w,h,paras,size=14,color=INK,bold=False,align='l',anchor='t',font=None,sp=1.15,after=0):
    tb=s.shapes.add_textbox(I(x),I(y),I(w),I(h)); tf=tb.text_frame; tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
    tf.vertical_anchor={'t':MSO_ANCHOR.TOP,'m':MSO_ANCHOR.MIDDLE,'b':MSO_ANCHOR.BOTTOM}[anchor]
    if isinstance(paras,str): paras=[paras]
    for i,pa in enumerate(paras):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment={'l':PP_ALIGN.LEFT,'c':PP_ALIGN.CENTER,'r':PP_ALIGN.RIGHT}[align]
        p.line_spacing=sp; p.space_after=Pt(after)
        runs=pa if isinstance(pa,list) else [(pa,{})]
        for t,o in runs:
            _run(p,t,o.get('size',size),o.get('color',color),o.get('bold',bold),o.get('font',font))
    return tb
def frame(s,n,title,source=''):
    rect(s,0.7,0.42,0.55,0.06,ACC)
    tx(s,0.7,0.62,11.9,1.0,title,size=26,color=NAVY,bold=True,font=LAT,sp=1.1)
    if source: tx(s,0.7,7.0,10.6,0.3,source,size=10,color=GREY)
    tx(s,11.9,7.0,0.75,0.3,f'{n}/{TOTAL}',size=10,color=GREY,align='r')
def new(): return prs.slides.add_slide(BLANK)


ILL='Illustrative data for demonstration'
# 1 cover
s=new(); rect(s,0,0,13.333,7.5,NAVY); rect(s,0.8,0.9,0.7,0.07,ACC)
tx(s,0.8,1.15,11,0.4,'BOARD STRATEGY REVIEW  ·  Q1 2026',size=16,color=SOFT)
tx(s,0.8,1.9,11.5,2.4,['Growth is intact,','but now depends on one channel'],size=48,color=WHITE,bold=True,sp=1.05)
tx(s,0.8,4.45,10.5,0.5,'Three facts, one decision, a 4-quarter plan',size=20,color=MID)
rect(s,0.8,5.6,11.7,0.02,C(0x2B,0x3F,0x52))
for i,(n,l) in enumerate([('+23%','revenue vs prior year'),('70%','of net new revenue from digital'),('$120M','diversification program')]):
    x=0.8+i*4.0
    tx(s,x,5.8,3.6,0.9,[[(n,{'size':40,'font':NUM,'color':WHITE if i<2 else C(0x5B,0x9B,0xFF),'bold':True})]])
    tx(s,x,6.6,3.7,0.35,l,size=13,color=SOFT)
# 2 exec summary
s=new(); frame(s,2,'Revenue is ahead of plan, yet 70% of the growth comes from a single channel',ILL)
cols=[('RESULT','+23%','revenue growth','Reached $4.2B, 5 points above guidance.'),
      ('RISK','70%','from one channel','Digital delivers most net new revenue; stores are flat.'),
      ('DECISION','$120M','over four quarters','Fund a channel-diversification program starting this month.')]
for i,(lab,big,unit,desc) in enumerate(cols):
    x=0.7+i*4.1; w=3.8
    rect(s,x,2.0,w,4.6,LIGHT); rect(s,x,2.0,w,0.07,ACC if i==2 else NAVY)
    tx(s,x+0.3,2.35,w-0.6,0.4,lab,size=14,color=ACC,bold=True)
    tx(s,x+0.3,2.85,w-0.6,1.1,[[(big,{'size':58,'font':NUM,'color':NAVY,'bold':True})]])
    tx(s,x+0.3,4.05,w-0.6,0.4,unit,size=16,color=GREY,bold=True)
    rect(s,x+0.3,4.6,0.5,0.03,MID)
    tx(s,x+0.3,4.8,w-0.6,1.6,desc,size=15,color=INK,sp=1.3)
# 3 native clustered column
s=new(); frame(s,3,'Digital doubled in two years while stores and wholesale stood still',ILL)
cd=CategoryChartData(); cd.categories=['Digital','Stores','Wholesale','Partners']
cd.add_series('2024',[1.2,1.6,0.9,0.3]); cd.add_series('2025',[2.1,1.6,0.9,0.4])
gf=s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,I(0.7),I(1.9),I(8.4),I(4.9),cd); ch=gf.chart
ch.has_title=False; ch.has_legend=True; ch.legend.include_in_layout=False
from pptx.enum.chart import XL_LEGEND_POSITION
ch.legend.position=XL_LEGEND_POSITION.TOP; ch.legend.font.size=Pt(13); ch.legend.font.name=LAT
va=ch.value_axis; va.visible=False; va.has_major_gridlines=False
ca=ch.category_axis; ca.tick_labels.font.size=Pt(14); ca.tick_labels.font.name=LAT; ca.format.line.color.rgb=MID
pl=ch.plots[0]; pl.gap_width=60; pl.overlap=-5
for k,col in enumerate([MID,ACC]):
    se=pl.series[k]; se.format.fill.solid(); se.format.fill.fore_color.rgb=col
    se.data_labels.show_value=True; se.data_labels.number_format='$0.0"B"'; se.data_labels.number_format_is_linked=False
    se.data_labels.font.size=Pt(13); se.data_labels.font.name=LAT; se.data_labels.position=XL_LABEL_POSITION.OUTSIDE_END
rect(s,9.6,1.9,3.05,4.7,LIGHT); rect(s,9.6,1.9,0.07,4.7,ACC)
tx(s,9.9,2.15,2.6,0.4,'Digital, 2 years',size=15,color=GREY,bold=True)
tx(s,9.9,2.65,2.7,1.0,[[('+75%',{'size':48,'font':NUM,'color':ACC,'bold':True})]])
tx(s,9.9,3.85,2.6,2.6,['$1.2B to $2.1B, while stores and wholesale did not move.','','Every other line is flat or small.'],size=15,color=INK,sp=1.3)
# 4 waterfall via stacked column (native)
s=new(); frame(s,4,'Price and volume added $40M; FX and costs gave back $20M',ILL)
cats=['Q1 base','Price','Volume','FX','Costs','Q1 total']
base=[0,100,115,132,120,0]; up=[0,15,25,0,0,0]; down=[0,0,0,8,12,0]; tot=[100,0,0,0,0,120]
cd=CategoryChartData(); cd.categories=cats
cd.add_series('base',base); cd.add_series('total',tot); cd.add_series('up',up); cd.add_series('down',down)
gf=s.shapes.add_chart(XL_CHART_TYPE.COLUMN_STACKED,I(0.7),I(1.8),I(8.6),I(5.0),cd); ch=gf.chart
ch.has_legend=False; ch.has_title=False
va=ch.value_axis; va.visible=False; va.has_major_gridlines=False; va.minimum_scale=0; va.maximum_scale=150
ca=ch.category_axis; ca.tick_labels.font.size=Pt(14); ca.tick_labels.font.name=LAT; ca.format.line.color.rgb=MID
pl=ch.plots[0]; pl.gap_width=40
cols_=[None,NAVY,C(0x16,0xA3,0x4A),C(0xDC,0x26,0x26)]
for k,se in enumerate(pl.series):
    if k==0: se.format.fill.background(); se.format.line.fill.background(); continue
    se.format.fill.solid(); se.format.fill.fore_color.rgb=cols_[k]
lbls={1:{0:'100',5:'120'},2:{1:'+15',2:'+25'},3:{3:'-8',4:'-12'}}
for k,m in lbls.items():
    for idx,t in m.items():
        dl=pl.series[k].points[idx].data_label; tf=dl.text_frame; tf.text=t
        r=tf.paragraphs[0].runs[0]; r.font.size=Pt(14); r.font.bold=True; r.font.color.rgb=WHITE; r.font.name=LAT
        dl.position=XL_LABEL_POSITION.CENTER
rect(s,9.7,1.9,2.95,4.7,NAVY)
tx(s,10.0,2.15,2.4,0.4,'Net change',size=15,color=SOFT,bold=True)
tx(s,10.0,2.65,2.6,1.0,[[('+$20M',{'size':40,'font':NUM,'color':WHITE,'bold':True})]])
tx(s,10.0,3.85,2.45,2.6,['Operations created $40M.','Currency and cost inflation took back half.'],size=15,color=MID,sp=1.3)
# 5 matrix
s=new(); frame(s,5,'Start with high-impact, low-effort moves; defer the expensive bets',ILL)
gx,gy,gw,gh=1.3,1.75,7.6,4.7
rect(s,gx,gy,gw/2,gh/2,TINT); rect(s,gx+gw/2,gy,gw/2,gh/2,LIGHT); rect(s,gx,gy+gh/2,gw/2,gh/2,LIGHT); rect(s,gx+gw/2,gy+gh/2,gw/2,gh/2,C(0xFA,0xFA,0xFB))
rect(s,gx+gw/2-0.005,gy,0.01,gh,MID); rect(s,gx,gy+gh/2-0.005,gw,0.01,MID)
for t,x,y,col in [('Do now',gx+0.15,gy+0.1,ACC),('Plan',gx+gw/2+0.15,gy+0.1,GREY),('Fill in',gx+0.15,gy+gh-0.4,GREY),('Defer',gx+gw/2+0.15,gy+gh-0.4,GREY)]:
    tx(s,x,y,2,0.3,t,size=14,color=col,bold=True)
tx(s,0.45,gy+gh/2-0.4,0.8,0.8,['Impact'],size=14,color=NAVY,bold=True,align='c',anchor='m')
tx(s,gx,gy+gh+0.08,gw,0.3,'Effort   low → high',size=14,color=NAVY,bold=True,align='c')
for t,fx,fy,hot in [('Process automation',.04,.84,1),('Pricing rules',.10,.62,1),('Market expansion',.56,.84,0),('M&A targets',.62,.62,0),('Minor fixes',.06,.32,0),('Legacy rewrite',.58,.28,0)]:
    w=0.15*len(t)+0.5; x=min(gx+fx*gw,gx+gw-w-0.1); y=gy+(1-fy)*gh-0.2
    rect(s,x,y,w,0.4,ACC if hot else WHITE,None if hot else SOFT)
    tx(s,x,y,w,0.4,t,size=14,color=WHITE if hot else INK,bold=bool(hot),align='c',anchor='m')
rect(s,9.4,1.75,3.25,4.7,NAVY)
tx(s,9.7,2.0,2.7,0.4,'Start here',size=16,color=WHITE,bold=True)
for i,(a,b) in enumerate([('Process automation','Quick win, funds the next step'),('Pricing rules','Margin without new spend')]):
    y=2.6+i*1.4
    tx(s,9.7,y,0.6,0.7,[[(str(i+1),{'size':30,'font':NUM,'color':C(0x5B,0x9B,0xFF),'bold':True})]])
    tx(s,10.3,y,2.3,0.4,a,size=15,color=WHITE,bold=True); tx(s,10.3,y+0.4,2.25,0.8,b,size=14,color=MID,sp=1.2)
# 6 gantt
s=new(); frame(s,6,'The program lands in four quarters, with the first decision gate at the end of Q2',ILL)
gx0=3.7; wk=(12.65-gx0)/4
for k,q in enumerate(['Q1','Q2','Q3','Q4']):
    x=gx0+k*wk; rect(s,x,1.85,wk-0.05,0.45,NAVY); tx(s,x,1.85,wk-0.05,0.45,q,size=15,color=WHITE,bold=True,align='c',anchor='m')
    rect(s,x+wk/2-0.005,2.35,0.01,3.9,LIGHT)
rows=[('Foundation',0,1,ACC),('Pilot in two regions',1,2,NAVY),('Scale rollout',2,4,NAVY),('Review and reset',3.2,4,SOFT)]
for r,(t,a,b,col) in enumerate(rows):
    y=2.6+r*0.9
    tx(s,0.7,y,2.9,0.55,t,size=16,color=INK,bold=True,anchor='m')
    rect(s,gx0+a*wk,y+0.05,(b-a)*wk-0.05,0.45,col)
mx=gx0+2.0*wk; rect(s,mx-0.015,2.35,0.03,3.9,ACC)
rect(s,mx-1.1,6.3,2.2,0.45,ACC); tx(s,mx-1.1,6.3,2.2,0.45,'Decision gate',size=14,color=WHITE,bold=True,align='c',anchor='m')
out=sys.argv[1] if len(sys.argv)>1 else 'freeform_showcase.pptx'
prs.save(out); full_cleanup(out); print('saved')
