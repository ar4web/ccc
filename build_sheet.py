"""Rebuild the supplied Basma workbook. Usage: python build_sheet.py [source.xlsx]
Uses one employee master and one live attendance register. BASMA_VERIFY=1
checks the model; BASMA_OPS_VERIFY=1 checks the consolidated operations layout.
Both run on isolated copies, never on live employee inputs.
LibreOffice recalculates a temporary copy; only calculated values are brought back
so the delivered Excel file retains its original charts, validation and layout.
"""
from pathlib import Path
import sys, subprocess, tempfile, urllib.request, re, math, zipfile, copy, datetime as dt, collections, os
from lxml import etree as ET
import openpyxl as xl
from openpyxl.utils import get_column_letter as L
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.hyperlink import Hyperlink
from openpyxl.formatting.rule import FormulaRule, DataBarRule
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart import DoughnutChart, BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from openpyxl.worksheet.pagebreak import Break
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.page import PageMargins
from openpyxl.chart.data_source import NumData, NumVal, AxDataSource, StrRef, StrData, StrVal
from openpyxl.utils.datetime import to_excel
ROOT=Path(__file__).resolve().parent
src=Path(sys.argv[1]) if len(sys.argv)>1 else Path('/tmp/basma_source/input.xlsx')
if not src.is_file():
 src.parent.mkdir(parents=True,exist_ok=True);urllib.request.urlretrieve('https://www.designarena.ai/u/67c0e925-f646-471f-99c1-9ecc06994181',src)
source_file_path=Path(src).resolve()
wb=xl.load_workbook(src);raw=xl.load_workbook(src,data_only=True); original=list(wb.sheetnames)
# Standard formulas replace imported dummy functions and orphan references.
for key in list(wb.defined_names): del wb.defined_names[key]
for s in list(wb):
 if s.title!='timesheet Samples ':wb.remove(s)
order=['Dashboard','Workload Status','Timesheet and Salarysheet','Timesheet_Input','Salarysheet','Timesheet','Main Dataset','Employees','Clients','Audit','Expense Manage','Control Center','Slicer Data','timesheet Samples ']
for n in order[:-1]:wb.create_sheet(n)
wb._sheets=[wb[n] for n in order];wb.active=0
wb.calculation=xl.workbook.properties.CalcProperties(calcMode='auto',fullCalcOnLoad=True,forceFullCalc=True,calcOnSave=True)
wb.properties.title='Workforce & payroll | Basmat Al-Mawared';wb.properties.description='Monthly payroll controls, source reconciliation and workforce reporting.'
N=425;E=355;C=45;X=255; start=dt.datetime(2026,9,1); source=wb['timesheet Samples ']
NAVY='163239'; DEEP='102B32'; TEAL='006A65'; AQUA='61D4BA'; GREEN='1D705B'; WHITE='FFFFFF'; INK='193940'; MUT='516C70'; MINT='E2F3EB'; PALE='F3F7F5'; SKY='EDF5FA'; BLUE='215E82'; GOLD='9B5717'; SAND='FCF0DE'; BORDER='D8E4E1'; ROSE='FBE7E1'; RED='9A3D29'; BG='F4F7F5'
MONEY='#,##0.00;[Red](#,##0.00);"—"';NUM='#,##0.##;[Red](#,##0.##);"—"';INT='#,##0;[Red](#,##0);"—"';PCT='0.0%;[Red](0.0%);"0.0%"';DATE='dd mmm yyyy';PER='mmm yyyy'
fills={};fonts={}; aligns={}
def fill(c):
 if c not in fills:fills[c]=PatternFill('solid',fgColor=c)
 return fills[c]
def font(c=INK,sz=10,b=False):
 k=c,sz,b
 if k not in fonts:fonts[k]=Font(name='Calibri',size=sz,bold=b,color=c)
 return fonts[k]
def align(h='left',ind=0,wrap=False):
 k=h,ind,wrap
 if k not in aligns:aligns[k]=Alignment(horizontal=h,vertical='center',indent=ind,wrap_text=wrap)
 return aligns[k]
def style(c,fc=INK,bg=WHITE,size=10,bold=False,h='left',ind=1,wrap=False):
 c.font=font(fc,size,bold);c.fill=fill(bg);c.alignment=align(h,ind,wrap);return c

def merge(s,r,c1,c2,v,fc=INK,bg=WHITE,size=10,bold=False,h='left',ind=1):
 if c2>c1:s.merge_cells(start_row=r,start_column=c1,end_row=r,end_column=c2)
 c=s.cell(r,c1,v);return style(c,fc,bg,size,bold,h,ind)

def nav(s,ref,sh,label=None):
 c=s[ref];c.value=label or sh;c.hyperlink=Hyperlink(ref=ref,location="'"+sh+"'!A1",display=label or sh);style(c,TEAL,MINT,10,True)

def setup(s):
 s.sheet_view.showGridLines=False;s.sheet_view.zoomScale=90;s.sheet_properties.pageSetUpPr.fitToPage=True;s.page_setup.orientation='landscape';s.page_setup.paperSize=s.PAPERSIZE_A3;s.page_setup.fitToWidth=1;s.page_setup.fitToHeight=0;s.page_margins=PageMargins(left=.32,right=.32,top=.45,bottom=.45,header=.2,footer=.2);s.oddFooter.left.text='BASMAT AL-MAWARED   /   CONFIDENTIAL';s.oddFooter.right.text='&P  /  &N';s.sheet_format.defaultRowHeight=18

def grid(sh,title,sub,heads,widths,end,inputs=(),freeze='A6',top=6,navto=None):
 s=wb[sh];setup(s);n=len(heads); merge(s,1,1,top,title,WHITE,NAVY,20,True);s.row_dimensions[1].height=39;merge(s,2,1,top,sub,MUT,PALE,11);s.row_dimensions[2].height=26
 nav(s,'A3','Dashboard','←  OVERVIEW');nav(s,'B3',navto or 'Control Center','OPEN  /  '+('GUIDE' if not navto else navto.upper()));s.row_dimensions[3].height=20
 for k,(h,wi) in enumerate(zip(heads,widths),1):
  s.column_dimensions[L(k)].width=wi;style(s.cell(5,k,h),WHITE,TEAL if k in inputs else NAVY,10,True,'left',1,True)
 s.row_dimensions[5].height=32;s.freeze_panes=freeze;s.auto_filter.ref=f'A5:{L(n)}{end}';s.print_title_rows='1:5';s.print_area=f'A1:{L(n)}{end}';s.sheet_properties.tabColor=TEAL if inputs else NAVY
 return s

def value(s,r,k,v,input=False,fmt=None,h=None):
 c=s.cell(r,k);c.value=v
 num=fmt is not None and fmt!='@';style(c,BLUE if input else INK,(SKY if r%2==0 else 'F5F9FB') if input else (WHITE if r%2==0 else PALE),10,False,h or ('right' if num else 'left'),1)
 if fmt:c.number_format=fmt
 return c

def dv(s,rng,typ,a,b=None,operator=None,msg='Choose a listed value.',blank=True):
 v=DataValidation(type=typ,formula1=a,formula2=b,operator=operator,allow_blank=blank,showErrorMessage=True,showInputMessage=True,errorTitle='Check this entry',error=msg,promptTitle='Input guide',prompt=msg);s.add_data_validation(v);v.add(rng);return v

def flag(s,rng,expr,bg=ROSE,fc=RED):s.conditional_formatting.add(rng,FormulaRule(formula=[expr],fill=fill(bg),font=font(fc,10,True)))
def hide(s,first,last):s.column_dimensions.group(first,last,hidden=True)
def rng(sh,col,end=N,begin=6):return f"'{sh}'!${col}${begin}:${col}${end}"
def nm(name,loc):wb.defined_names.add(DefinedName(name,attr_text=loc))
def arr(lst):return '"'+','.join(lst)+'"'
def cleanid(x):return str(int(x)) if isinstance(x,(float,int)) else str(x).strip() if x is not None else None
def datev(v):
 if isinstance(v,str):
  try:return dt.datetime.strptime(v,'%d-%m-%Y')
  except ValueError:pass
 return v

def ordered(vs):return sorted(set(v for v in vs if v is not None),key=str)
def cfbar(s,rg,color):s.conditional_formatting.add(rg,DataBarRule(start_type='min',end_type='max',color=color,showValue=True))
emp=raw['Employees'];rc=raw['Clients'];reg=raw['Timesheet and Salarysheet'];inp=raw['Timesheet_Input'];ex=raw['Expense Manage']; ep=[list(r) for r in emp.iter_rows(min_row=3,max_row=205,max_col=21,values_only=True)];cards=[list(r) for r in rc.iter_rows(min_row=2,max_row=17,max_col=11,values_only=True)];extra=ordered([r[6] for r in ep if r[6] not in [c[0] for c in cards]])
# Control center: settings, navigation, source lineage and operating policy.
g=wb['Control Center'];setup(g);
for col,wi in zip('ABCDEFGH',[31,23,20,20,20,20,19,19]):g.column_dimensions[col].width=wi
merge(g,1,1,8,'Control center',WHITE,NAVY,24,True);g.row_dimensions[1].height=44
merge(g,2,1,8,'OPERATING GUIDE    /    Set the reporting scope on Dashboard. Blue cells are editable; white and mint cells calculate.',MUT,PALE,11);g.row_dimensions[2].height=27
nav(g,'A3','Dashboard','←  OPEN DASHBOARD');nav(g,'B3','Timesheet_Input','ATTENDANCE ENTRY');nav(g,'C3','Clients','RATE CARDS');nav(g,'D3','Audit','AUDIT QUEUE')
def band(s,r,txt,c1=1,c2=8):merge(s,r,c1,c2,txt,WHITE,TEAL,11,True);s.row_dimensions[r].height=23
band(g,5,'01    /    Operating settings');g['A6']='Setting';g['B6']='Input';merge(g,6,3,8,'How the value is used',WHITE,NAVY,10,True)
for k in 'AB':style(g[k+'6'],WHITE,NAVY,10,True)
settings=[('Iqama alert window',30,'Days from the compliance date. Negative days are expired; zero means due on the reference date.'),('Insurance alert window',30,'Days from the compliance date. Insurance and Iqama checks are independent.'),('Expense approval limit',1000,'SAR • claims above this source threshold are marked for HR review.'),('Compliance date','=TODAY()','Defaults to today, not the payroll month. Enter a fixed date for a repeatable compliance review.'),('Weekly premium cycle',7,'Days between premium slots. The source PLASTIC rule uses days 1, 8, 15, 22 and 29.'),('First premium day',1,'1 = first calendar day of the payroll period. Used only by Cycle premium clients.'),('Half-day absence charge',0,'0% preserves the source rule: 0.5P adds half a present day but no absence charge. 100% charges the missing half.'),('Annual leave allowance',30,'Reference only • source assumption; leave transactions were not included.'),('Sick leave allowance',10,'Reference only • source assumption; no balance is inferred.'),('Attrition threshold',.05,'Reference only • exit dates were not supplied; historical attrition is not estimated.'),('Default OT multiplier',1.5,'Reference for new rate cards. Enter an explicit multiplier on every payable client.'),('Currency','SAR','Saudi riyal • all pay and expense amounts are in SAR; two-decimal settlement.')]
for r,(label,v,note) in enumerate(settings,7):
 value(g,r,1,label);value(g,r,2,v,True,DATE if r==10 else PCT if r in [13,16] else None);merge(g,r,3,8,note,MUT,WHITE);g.row_dimensions[r].height=23
for r,lo,hi in [(7,0,365),(8,0,365),(11,1,31),(12,1,31)]:dv(g,f'B{r}','whole',lo,hi,'between','Enter a whole number in range.')
dv(g,'B13','decimal',0,1,'between','Enter 0% to 100%.');dv(g,'B10','date','DATE(2000,1,1)','DATE(2100,12,31)','between','Enter a valid reference date.');dv(g,'B9','decimal',0,1000000,'between','Enter a nonnegative SAR limit.')
def guide(r,label,txt):
 value(g,r,1,label);c=merge(g,r,2,8,txt,INK,'F3F7F5' if r%2 else WHITE,11);c.alignment=align('left',1,True);g.row_dimensions[r].height=33
band(g,20,'02    /    Monthly workflow')
for r,txt in enumerate([
 ('01  •  Maintain master data','Edit employee details only in Main Dataset. Employees is a linked directory; payroll and reporting use that directory. Clients holds the contractual pay rules.'),('02  •  Enter one line per ID/month','Timesheet_Input is the only attendance source. Type a month-start date, submitted site, exact employee name and IQAMA. Keep prior months on their own rows; changing the dashboard period never relabels attendance.'),('03  •  Record attendance','Use P, A, OFF, 0.5P or a numeric entry. Blank means missing, not absence. Use only actual calendar days. Monthly-base numeric positives mean one attended day; zero means absent.'),('04  •  Review payroll','Salarysheet separates accrued net, held net and ready net. Held values are estimates, not deductions or approvals. A ready result means spreadsheet checks passed; authorize payment separately.'),('05  •  Resolve issues','Use Audit to find the source row and first blocking reason. Check master name and site before changing identity data. Other issues may surface when the first is resolved.'),('06  •  Explore the dashboard','Choose the month and filters. Worksite means submitted site for payroll, current site for the roster. Sponsor, contract, nationality and status come from the current master. Unknown IDs are Unmatched.'),('07  •  Read the client statement','Timesheet and Salarysheet follows all overview filters and adds a Ready / Hold selector. Timesheet shows the current roster and whether each IQAMA has an entry for the chosen month.'),('08  •  Prepare the next period','Add a new input row or save a separate workbook for the next period. 420 payroll lines, 350 master rows, 40 site records and 250 expense lines are prewired. Do not enter outside the prepared ranges.'),],22):guide(r,*txt)
band(g,31,'03    /    Calculation rules & release checks')
for r,v in enumerate([
 ('Monthly base','Employee monthly salary is used when numeric and positive. The client salary may estimate a missing or invalid base, but that row stays on hold. No silent SAR 1,200 default.'),('Rates and deductions','Client daily rate = ROUND(client monthly salary ÷ contractual days, 2); hourly rate = ROUND(daily rate ÷ shift hours, 2). Absence charge = absent equivalents × daily rate × unpaid share.'),('Absence treatment','Unpaid / Half Pay / Full Pay charge 100% / 50% / 0% respectively. OFF is logged but never an absence. Other SAR deductions are input per period, not attached permanently to a name. Total deductions are capped at base pay.'),('Rate-card alternative','Rate card basis uses the client format: P/A Tally and Direct Days earn day units; 8-Hour Daily and Hourly Total earn regular hours; Mixed follows pay method. P = one shift, 0.5P = half a shift. Enter regular hours only; approved OT is separate.'),('Overtime precedence','No OT or a zero multiplier pays zero. Otherwise an explicit OT amount (including zero) overrides calculations. Approved hours uses hours × client hourly rate × multiplier. Cycle premium uses qualifying worked slots × the site premium.'),('Identity controls','One IQAMA per month, one master record per IQAMA. Names and sites must match after normalizing spaces/case; substring matches are not accepted. A different worksite must be resolved, not silently reassigned.'),('Attendance controls','Release requires all calendar days logged with recognized codes, no entries after month-end, valid client rules and nonnegative adjustments. Rate-card day numbers must be 0 or 1; regular hours must not exceed a shift. Half days use 0.5P.'),('Calendar conflict','The original September register contains two columns dated 30 Sep. Only the first is counted. The second is retained in Input column AV. If the entries differ, review the source and set AX to Confirmed before release.'),('Roster versus history','A month filter selects transactions, not a reconstructed historical roster. Master status and compliance use the current master. The roster includes unsubmitted employees; they do not create payroll or absence entries.'),('Historical limitation','Rates and employee data are live: corrections affect all periods in this copy. Save a dated, approved copy for each closed period. This workbook does not establish statutory compliance or a payment authorization.'),('Unobserved metrics','No exit-date or leave-transaction register was supplied. Attrition, retention and leave taken are not fabricated; use work status and dated attendance instead.'),('Gross-to-net bridge','Gross = base + overtime. Net estimate = gross − capped deductions. Ready net + held net = net estimate. Missing rates can make held estimates incomplete; review the number of held records, not only the held value.')],33):guide(r,*v)
band(g,46,'04    /    Source reconciliation')
for r,v in enumerate([('Canonical register','Imported the 73 populated rows from Timesheet and Salarysheet for September 2026. The 50 populated Timesheet_Input rows overlap that register; their daily entries were compared and not double-counted.'),('Adjustments carried forward','Fixed OT amounts from source AP and row-specific extra deductions from AQ were carried as period inputs. PLASTIC retains the source weekly SAR 65 premium rather than inferring another OT rule from a rate-card note.'),('Duplicate day retained','All second 30 September entries are retained for inspection. The original first-30th entry is used unless you correct it. Reconciliation checks use actual calendar days, including leap years.'),('Master and supporting data','All 203 employee records, 16 contractual client cards and source expense rows are retained. Additional roster-only sites have no invented rates. October samples remain examples, not payroll submissions.'),('Original versus active policy','Current production rows use the source monthly-base salary approach. Rate card basis is an explicit alternative for new calculations. Confirm contractual rules, imported overrides and deductions before closing payroll.'),('Navigation','All original worksheet roles remain. Workload Status is a live site scorecard; Audit lists holds and employment exceptions. Slicer Data is a hidden formula helper, not an input surface.')],48):guide(r,*v)
g.freeze_panes='C7';g.print_area='A1:H54';g.sheet_properties.tabColor=BLUE
# Reference lists and formula-fed worksite list.
ls=wb['Slicer Data'];setup(ls);ls.sheet_state='hidden'
lookupLists={'Sites':[None]*40,'Sponsors':ordered(r[7] for r in ep)+['Unmatched'],'Nations':ordered([r[12] for r in ep]+[raw['Control Center'].cell(r,1).value for r in range(116,126)])+['Unmatched'],'Contracts':ordered([r[16] for r in ep]+[raw['Control Center'].cell(r,5).value for r in range(116,120)])+['Unmatched'],'Statuses':['Active','On Leave','Suspended','Terminated','Unmatched'],'Sort':['Ready net','Net estimate','Absence days','Attendance rate'],'Periods':[dt.datetime(y,m,1) for y in range(2025,2032) for m in range(1,13)]}
for j,(name,items) in enumerate(lookupLists.items(),1):
 ls.cell(1,j,name);allflag=name not in ['Sort','Periods'];fst=3 if allflag else 2
 if allflag:ls.cell(2,j,'All')
 for r,v in enumerate(items,fst):
  ls.cell(r,j,f'=IF(\'Clients\'!A{r+3}="","",\'Clients\'!A{r+3})' if name=='Sites' else v)
  if name=='Periods':ls.cell(r,j).number_format=PER
 nm(name,f"'Slicer Data'!${L(j)}$2:${L(j)}${fst+len(items)-1}")
# Dashboard selection references. All filters are applied in the same helper flags.
for name,addr in [('ViewMonth','B9'),('ViewSite','F9'),('ViewSponsor','J9'),('ViewStatus','N9'),('ViewNation','B12'),('ViewContract','F12'),('ViewSort','J12')]:nm(name,"'Dashboard'!$"+re.sub(r'(\D+)(\d+)',r'\1$\2',addr))
# Editable employee master: keep the original order and information, normalize real dates and IDs.
heads=[emp.cell(2,j).value for j in range(1,22)]+['Iqama status','Insurance status','Master check','In scope','ID key','Employment audit','Audit no.']
employee_widths=[13,32,18,18,15,18,22,29,21,16,18,19,20,18,30,17,18,18,12,25,22,20,20,20,12,18,22,12]
master=grid('Main Dataset','Employee master data','EMPLOYEE SOURCE    /    Edit blue fields here. Changes flow through Employees to payroll checks, roster totals and Iqama expiry cards.',heads[:21],employee_widths[:21],E,set(range(1,22))-{5},'D6',7,'Employees')
s=grid('Employees','Employee directory','LINKED DIRECTORY    /    All employee attributes come from Main Dataset. Search below; use Open / Main Dataset to edit a record.',heads,employee_widths,E,(),'D6',6,'Main Dataset')
for r in range(6,E+1):
 v=ep[r-6] if r<209 else [None]*21
 for j,a in enumerate(v,1):
  if j==5:continue
  if j in [4,6,11,18]:a=datev(a)
  if j==3:a=cleanid(a)
  fmt=DATE if j in (4,6,11,18) else MONEY if j in (16,19) else '@' if j in (3,12,15) else None
  value(master,r,j,a,True,fmt)
  ref=f"'Main Dataset'!{L(j)}{r}"
  value(s,r,j,f'=IF(ISBLANK({ref}),"",{ref})',fmt=fmt)
 value(master,r,5,f'=IF(OR(B{r}="",NOT(ISNUMBER(D{r})),N(D{r})<=0,NOT(ISNUMBER(\'Control Center\'!$B$10))),"",INT(D{r})-INT(\'Control Center\'!$B$10))',fmt=INT)
 for j,f in {5:f"'Main Dataset'!E{r}",22:f'IF(B{r}="","",IF(NOT(ISNUMBER(E{r})),"Missing date",IF(E{r}<0,"Expired",IF(E{r}<=\'Control Center\'!$B$7,"Due soon","Current"))))',23:f'IF(B{r}="","",IF(NOT(ISNUMBER(R{r})),"Missing date",IF(R{r}<\'Control Center\'!$B$10,"Expired",IF(R{r}-\'Control Center\'!$B$10<=\'Control Center\'!$B$8,"Due soon","Current"))))',24:f'IF(B{r}="","",IF(C{r}="","ID missing",IF(COUNTIF($Z$6:$Z${E},Z{r})>1,"Duplicate ID",IF(OR(NOT(ISNUMBER(P{r})),P{r}<=0),"Check salary","OK"))))',25:f'IF(B{r}="",0,--AND(OR(ViewSite="All",G{r}=ViewSite),OR(ViewSponsor="All",H{r}=ViewSponsor),OR(ViewStatus="All",J{r}=ViewStatus),OR(ViewNation="All",M{r}=ViewNation),OR(ViewContract="All",Q{r}=ViewContract)))',26:f'IF(C{r}="","",TRIM(C{r}&""))',27:f'IF(B{r}="","",IF(OR(ISNUMBER(SEARCH("runaway",G{r}&" "&J{r}&" "&U{r})),ISNUMBER(SEARCH("left",G{r}&" "&J{r}&" "&U{r})),ISNUMBER(SEARCH("not interested",G{r}&" "&J{r}&" "&U{r})),ISNUMBER(SEARCH("resigned",G{r}&" "&J{r}&" "&U{r})),ISNUMBER(SEARCH("terminated",G{r}&" "&J{r}&" "&U{r}))),"Review",""))',28:f'IF(AND(AA{r}="Review",Y{r}=1),COUNT($AB$5:AB{r-1})+1,"")'}.items():value(s,r,j,'='+f,fmt=INT if j==5 else None)
hide(s,'Y','AB');flag(s,f'V6:W{E}','V6="Expired"');flag(s,f'V6:W{E}','V6="Due soon"','FFF3C4','856000');flag(s,f'X6:X{E}','AND($X6<>"OK",$X6<>"")');dv(master,f'J6:J{E}','list',arr(['Active','On Leave','Suspended','Terminated']));dv(master,f'P6:P{E}','decimal',0,1e8,'between','Enter a nonnegative numeric salary.');nm('ClientNames',rng('Clients','A',C))
# Protect the formula view while leaving its existing lookup box and filters available.
s.protection.sheet=True;s.protection.autoFilter=False
s['D3'].protection=xl.styles.Protection(locked=False)
for k in ('D','R'):dv(master,f'{k}6:{k}{E}','date','DATE(1900,1,1)','DATE(2199,12,31)','between','Enter a real expiry date. Leave blank if unknown.')
# Client rules, including source-only assignments with missing contracts visibly unconfigured.
heads=['Client / worksite','Monthly rate','Period days','Shift hours','Daily rate','Hourly rate','Timesheet format','Absence policy','Pay method','OT multiplier','Source policy note','Pay basis','OT rule','Cycle premium','Rate check']
s=grid('Clients','Client rate cards','POLICY MATRIX    /    SAR rates are contractual inputs. Roster-only sites stay unconfigured until a contract is entered.',heads,[23,18,14,14,16,16,20,18,18,16,84,19,21,18,22],C,{1,2,3,4,7,8,9,10,11,12,13,14},'B6',7,'Workload Status')
for r in range(6,C+1):
 vals=cards[r-6][:] if r<22 else [extra[r-22]]+[None]*10 if r-22<len(extra) else [None]*11
 if vals[0] and r>=22:vals[10]='Roster assignment only • no rate card in source.'
 vals+=['Monthly base' if r<22 else None,('Cycle premium' if vals[0]=='PLASTIC' else 'Approved hours' if vals[9] else 'No OT') if r<22 else None,65 if vals[0]=='PLASTIC' else None]
 for j,v in enumerate(vals,1):
  if j in [5,6]:continue
  value(s,r,j,v,True,MONEY if j in [2,14] else '0.00"×"' if j==10 else None)
 for j,f in {5:f'IF(OR(A{r}="",NOT(ISNUMBER(B{r})),NOT(ISNUMBER(C{r})),C{r}<=0),"",ROUND(B{r}/C{r},2))',6:f'IF(OR(NOT(ISNUMBER(E{r})),NOT(ISNUMBER(D{r})),D{r}<=0),"",ROUND(E{r}/D{r},2))',15:f'IF(A{r}="","",IF(COUNTIF($A$6:$A${C},A{r})>1,"Duplicate client",IF(AND(ISNUMBER(B{r}),B{r}>0,ISNUMBER(C{r}),C{r}>=1,C{r}<=31,INT(N(C{r}))=C{r},ISNUMBER(D{r}),D{r}>0,D{r}<=24,ISNUMBER(J{r}),J{r}>=0,OR(H{r}="Unpaid",H{r}="Half Pay",H{r}="Full Pay"),OR(L{r}="Monthly base",L{r}="Rate card"),OR(G{r}="P/A Tally",G{r}="Direct Days",G{r}="8-Hour Daily",G{r}="Hourly Total",G{r}="Mixed"),OR(I{r}="Hourly Rate",I{r}="Daily Rate"),OR(M{r}="No OT",M{r}="Approved hours",AND(M{r}="Cycle premium",ISNUMBER(N{r}),N{r}>0))),"Configured","Needs rates")))'}.items():value(s,r,j,'='+f,fmt=MONEY if j<7 else None)
for col,opts in [('G',['P/A Tally','Direct Days','8-Hour Daily','Hourly Total','Mixed']),('H',['Unpaid','Half Pay','Full Pay']),('I',['Daily Rate','Hourly Rate']),('L',['Monthly base','Rate card']),('M',['Approved hours','Cycle premium','No OT'])]:dv(s,f'{col}6:{col}{C}','list',arr(opts))
for col,lo,hi,ty in [('B',.01,1e8,'decimal'),('C',1,31,'whole'),('D',.01,24,'decimal'),('J',0,10,'decimal'),('N',0,1e6,'decimal')]:dv(s,f'{col}6:{col}{C}',ty,lo,hi,'between','Enter a valid positive rate or permitted zero.')
flag(s,f'O6:O{C}','AND(O6<>"",O6<>"Configured")',SAND,GOLD)
# Attendance entry: one period-tagged record per worker. The surplus date is kept as evidence.
heads=['Period','Submitted client','Submitted name','IQAMA','Identity check','Master name','Master client']+[f'Day {j:02}' for j in range(1,32)]+['Present eq.','Off days','Absent days','Days logged','Entry check','OT hours','OT override','Other deduction','Import / adjustment note','Extra 30 Sep','Source row','Date review','Master index','Client index','Sponsor','Nationality','Contract','Work status','In scope','Month days','Basis','Format','Shift','Hour basis','Bad codes','Master salary','ID key','Has input','Source conflict']
s=grid('Timesheet_Input','Attendance entry','MONTHLY REGISTER    /    Enter one row per IQAMA and month. P · A · OFF · 0.5P · numeric time. Blank days are never counted as absences.',heads,[16,23,33,17,23,34,22]+[9]*31+[15,13,15,14,27,13,16,18,47,16,13,17]+[15]*17,N,{1,2,3,4,*range(8,39),44,45,46,47,50},'F6',7,'Salarysheet')
for r in range(6,N+1):
 populated=r<=78;old=reg[r] if populated else []; site=old[0].value if old else None;name=old[2].value if old else None
 vals={1:start,2:site,3:name,4:cleanid(old[3].value)} if old else {};ded=0;fixed=None;conf=False
 if old:
  for j in range(30):vals[8+j]=old[5+j].value
  fixed=old[41].value if r>=56 else None
  if r>=56:
   # The source AQ formula stores a literal per-record adjustment.
   # Source values with the explicit 80.23 addend are detected from the formula by the reader below.
   if r==56:sourceF=xl.load_workbook(src,read_only=True)['Timesheet and Salarysheet'];srcAQ={row[0].row:row[0].value for row in sourceF.iter_rows(min_row=56,max_row=78,min_col=43,max_col=43)}
   m=re.search(r'\+\s*([\d.]+)\)\s*,\s*2\)\)',srcAQ[r]);assert m,(r,srcAQ[r]);ded=float(m.group(1))
  else:ded=100 if name=='AL AMIN KHAN' else 50 if name in ['NAYAN DAS','MD AKTER HOSSAIN','MD ATAUR RAHMAN','MATIUR RAHMAN'] else 0
  conf=old[34].value!=old[35].value
  vals.update({45:fixed,46:ded if ded else None,47:('Source OT amount; ' if fixed is not None else '')+('source deduction; ' if ded else '')+('two 30 Sep entries differ' if conf else 'September register'),48:old[35].value,49:r,50:'Review' if conf else None,67:1 if conf else 0})
 for j in [1,2,3,4,*range(8,39),44,45,46,47,48,49,50,67]:value(s,r,j,vals.get(j),j not in (48,49,67),PER if j==1 else '@' if j==4 else MONEY if j in [45,46] else None,'center' if j in range(8,39) else None)
 days=f'$H{r}:INDEX($H{r}:$AL{r},1,MAX(1,$BF{r}))';allD=f'$H{r}:$AL{r}'
 def ei(col,default='"Unmatched"'):return f'IF(AY{r}=0,{default},INDEX({rng("Employees",col,E)},AY{r}))'
 def ci(col,default='""'):return f'IF(AZ{r}=0,{default},INDEX({rng("Clients",col,C)},AZ{r}))'
 normal=lambda c:f'TRIM(SUBSTITUTE({c}{r},CHAR(160)," "))'
 fs={66:f'--(COUNTA(A{r}:D{r},H{r}:AL{r},AR{r}:AT{r})>0)',65:f'IF(D{r}="","",TRIM(D{r}&""))',51:f'IF(BM{r}="",0,IFERROR(MATCH(BM{r},{rng("Employees","Z",E)},0),0))',52:f'IF(B{r}="",0,IFERROR(MATCH(B{r},{rng("Clients","A",C)},0),0))',58:f'IFERROR(IF(AND(ISNUMBER(A{r}),A{r}=INT(A{r}),DAY(A{r})=1),DAY(DATE(YEAR(A{r}),MONTH(A{r})+1,0)),0),0)',5:f'IF(BN{r}=0,"",IF(D{r}="","ID missing",IF(COUNTIFS($BM$6:$BM${N},BM{r},$A$6:$A${N},A{r})>1,"Duplicate entry",IF(AY{r}=0,"ID not found",IF(COUNTIF({rng("Employees","Z",E)},BM{r})>1,"Duplicate master ID",IF(AND({normal("C")}<>{normal("F")},{normal("B")}<>{normal("G")}),"Name & site mismatch",IF(OR(C{r}="",{normal("C")}<>{normal("F")}),"Name mismatch",IF(OR(B{r}="",{normal("B")}<>{normal("G")} ),"Site mismatch","Matched"))))))))',6:f'IF(BN{r}=0,"",{ei("B",chr(34)+chr(34))})',7:f'IF(BN{r}=0,"",{ei("G",chr(34)+chr(34))})',39:f'IF(BN{r}=0,"",COUNTIF({days},"P")+0.5*COUNTIF({days},"0.5P")+SUMPRODUCT(--ISNUMBER({days}),--({days}>0)))',40:f'IF(BN{r}=0,"",COUNTIF({days},"OFF"))',41:f'IF(BN{r}=0,"",COUNTIF({days},"A")+SUMPRODUCT(--ISNUMBER({days}),--({days}=0)))',42:f'IF(BN{r}=0,"",COUNTA({days}))',59:ci('L'),60:ci('G'),61:ci('D','0'),62:f'IF(BN{r}=0,0,--OR(BH{r}="8-Hour Daily",BH{r}="Hourly Total",AND(BH{r}="Mixed",{ci("I")}="Hourly Rate")))',63:f'IF(BN{r}=0,0,COUNTA({days})-COUNT({days})-COUNTIF({days},"P")-COUNTIF({days},"OFF")-COUNTIF({days},"A")-COUNTIF({days},"0.5P")+COUNTIF({days},"<0")+COUNTIF({days},">24")+IF(BG{r}="Rate card",IF(BJ{r}=1,COUNTIF({days},">"&BI{r}),COUNT({days})-SUMPRODUCT(--ISNUMBER({days}),--({days}=0))-COUNTIF({days},1)),0))',64:ei('P','0'),57:f'IF(BN{r}=0,0,--AND(A{r}=ViewMonth,OR(ViewSite="All",B{r}=ViewSite),OR(ViewSponsor="All",BA{r}=ViewSponsor),OR(ViewNation="All",BB{r}=ViewNation),OR(ViewContract="All",BC{r}=ViewContract),OR(ViewStatus="All",BD{r}=ViewStatus)))'}
 for j,col in zip(range(53,57),'HMQJ'):fs[j]=f'IF(BN{r}=0,"",{ei(col)})'
 checks=[(f'BF{r}=0','Invalid period'),(f'BK{r}<>0','Invalid day code'),(f'COUNTA({allD})>AP{r}','Outside month'),(f'AP{r}<BF{r}','Incomplete days'),(f'AND(BO{r}=1,AX{r}<>"Confirmed")','Confirm 30 Sep'),(f'OR(COUNTA(AR{r}:AT{r})<>COUNT(AR{r}:AT{r}),MIN(AR{r}:AT{r})<0)','Check adjustment'),(f'AND(OR({ci("M")}="No OT",{ci("J","0")}=0),SUM(AR{r}:AS{r})>0)','OT not allowed'),(f'AND({ci("M")}="Cycle premium",N(AR{r})>0)','OT hours unused')]
 body='"Complete"'
 for test,label in reversed(checks):body=f'IF({test},"{label}",{body})'
 fs[43]=f'IF(BN{r}=0,"",{body})'
 for j,f in fs.items():value(s,r,j,'='+f,fmt=NUM if j in range(39,43) else None)
# Helpers can be unhidden for inspection, but are not data entry fields.
hide(s,'AY','BO');s.auto_filter.ref=f'A5:AX{N}';s.print_area='A1:AQ78'
dv(s,f'A6:A{N}','custom','AND(ISNUMBER(A6),A6=INT(A6),DAY(A6)=1)','',msg='Use a real month-start date, for example 1 Sep 2026.');dv(s,f'B6:B{N}','list','ClientNames')
dv(s,f'H6:AL{N}','custom','OR(H6="",AND(COLUMN()-7<=$BF6,OR(H6="P",H6="A",H6="OFF",H6="0.5P",AND(ISNUMBER(H6),H6>=0,H6<=24))))',msg='Within the month: P, A, OFF, 0.5P, or 0–24 numeric units.')
dv(s,f'AR6:AT{N}','decimal',0,1000000,'between','Enter a nonnegative number; leave the OT override blank to calculate it.');dv(s,f'AX6:AX{N}','list',arr(['Review','Confirmed']));flag(s,f'E6:E{N}','AND($E6<>"",$E6<>"Matched")');flag(s,f'AQ6:AQ{N}','AND($AQ6<>"",$AQ6<>"Complete")',SAND,GOLD);flag(s,f'H6:AL{N}','AND($BN6=1,COLUMN()-7>$BF6)','E1E7E6',MUT);flag(s,f'H6:AL{N}','OR(H6="A",AND(ISNUMBER(H6),H6=0))',ROSE,RED);flag(s,f'H6:AL{N}','H6="OFF"',PALE,MUT)
# Payroll ledger; one-to-one with the input sheet regardless of dashboard filters.
heads=['Period','Client / worksite','Employee name','IQAMA','Identity check','Base pay','Absence charge¹','Other deduction¹','Overtime','Gross pay','Deductions','Net estimate','Held net','Ready net','Release','First review item','Logged days','Present eq.','Absent days','Off days','OT hours used','In scope','Report no.','Hold no.','Sort value','Sort rank','Daily rate','Hourly rate','Absence share','Pay basis','OT rule','OT multiplier','Premium','Payable hours','Rate check']
s=grid('Salarysheet','Payroll ledger','SETTLEMENT    /    SAR • Gross − deductions = net estimate = ready + held. ¹ Before the combined deduction cap. This ledger contains all periods.',heads,[16,24,34,18,24,17,19,20,16,17,17,18,17,17,14,25,16,15,15,13,17]+[15]*14,N,(), 'F6',7,'Timesheet_Input')
for r in range(6,N+1):
 t="'Timesheet_Input'!";a=lambda c:f'{t}{c}{r}';blank=f'{a("BN")}=0';idx=a('AZ');days=f'{a("H")}:INDEX({a("H")}:{a("AL")},1,MAX(1,{a("BF")}))';cl=lambda c,defv='0':f'IF({idx}=0,{defv},INDEX({rng("Clients",c,C)},{idx}))'
 fs={1:f'IF({a("A")}="","",{a("A")})',2:f'IF({a("B")}="","",{a("B")})',3:f'IF({a("C")}="","",{a("C")})',4:f'IF({a("D")}="","",{a("D")}&"")',5:f'IF({blank},"",{a("E")})'}
 inner={27:cl('E'),28:cl('F'),29:f'IF({cl("H")}="Full Pay",0,IF({cl("H")}="Half Pay",0.5,1))',30:a('BG'),31:cl('M','""'),32:cl('J'),33:cl('N'),34:f'IF({a("BJ")}=1,SUM({days})+(COUNTIF({days},"P")+0.5*COUNTIF({days},"0.5P"))*N({a("BI")}),0)',35:cl('O','"No rate card"'),6:f'IF(AD{r}="Rate card",ROUND(IF({a("BJ")}=1,IF({a("BH")}="8-Hour Daily",IFERROR(AH{r}/{a("BI")}*N(AA{r}),0),AH{r}*N(AB{r})),R{r}*N(AA{r}))+(1-AC{r})*S{r}*N(AA{r}),2),IF(AND(ISNUMBER({a("BL")}),N({a("BL")})>0),{a("BL")},N({cl("B")})))',7:f'IF(AD{r}="Rate card",0,ROUND((S{r}+COUNTIF({days},"0.5P")*0.5*\'Control Center\'!$B$13)*N(AA{r})*AC{r},2))',8:f'MAX(0,N({a("AT")}))',21:f'IF(AND(AE{r}="Approved hours",AF{r}>0,ISBLANK({a("AS")})),MAX(0,N({a("AR")})),0)',10:f'ROUND(F{r}+I{r},2)',11:f'ROUND(MIN(F{r},G{r}+H{r}),2)',12:f'ROUND(J{r}-K{r},2)',13:f'IF(O{r}="Hold",L{r},0)',14:f'IF(O{r}="Ready",L{r},0)',15:f'IF(P{r}="Checks passed","Ready","Hold")',16:f'IF(E{r}<>"Matched",E{r},IF(AI{r}<>"Configured","Check client rates",IF(AND(AD{r}="Monthly base",OR(NOT(ISNUMBER({a("BL")})),N({a("BL")})<=0)),"Check master salary",IF({a("AQ")}<>"Complete",{a("AQ")},"Checks passed"))))',17:a('AP'),18:a('AM'),19:a('AO'),20:a('AN'),22:a('BE'),23:f'IF(AND(V{r}=1,OR(\'Timesheet and Salarysheet\'!$D$3="All",\'Timesheet and Salarysheet\'!$D$3=O{r})),COUNT($W$5:W{r-1})+1,"")',24:f'IF(AND(V{r}=1,O{r}="Hold"),COUNT($X$5:X{r-1})+1,"")',25:f'IF(V{r}<>1,"",IF(ViewSort="Ready net",N{r},IF(ViewSort="Net estimate",L{r},IF(ViewSort="Absence days",S{r},IFERROR(R{r}/(R{r}+S{r}),0)))))',26:f'IF(ISNUMBER(Y{r}),RANK(Y{r},$Y$6:$Y${N},0)+COUNTIF($Y$6:Y{r},Y{r})-1,"")'}
 # The premium is paid only on an eligible worked slot, not on OFF text or a blank.
 allD=f'{a("H")}:{a("AL")}';dn=f'(COLUMN({allD})-7)';slot=f'SUMPRODUCT(--({dn}<={a("BF")}),--({dn}>=\'Control Center\'!$B$12),--(MOD({dn}-\'Control Center\'!$B$12,MAX(1,\'Control Center\'!$B$11))=0),(ISNUMBER({allD})*({allD}>0)+({allD}="P")+({allD}="0.5P")))'
 inner[9]=f'IF(OR(AE{r}="No OT",N(AF{r})=0),0,IF(ISNUMBER({a("AS")}),ROUND(MAX(0,{a("AS")}),2),IF(AE{r}="Cycle premium",ROUND({slot}*N(AG{r}),2),ROUND(U{r}*N(AB{r})*N(AF{r}),2))))'
 for j,f in inner.items():fs[j]=f'IF({blank},"",{f})'
 for j,f in fs.items():
  value(s,r,j,'='+f,fmt=PER if j==1 else MONEY if 6<=j<=14 else NUM if 17<=j<=21 else None)
hide(s,'V','AI');s.auto_filter.ref=f'A5:U{N}';s.print_area='A1:U78';flag(s,f'O6:O{N}','O6="Hold"');flag(s,f'O6:O{N}','O6="Ready"',MINT,GREEN)
# Roster-by-period workload register. No unstated absence days for missing entries.
heads=['Employee','IQAMA','Current site','Sponsor','Work status','Contract','Period entries','Submission','Present eq.','Absent days','Off days','Ready net','Held net','Iqama status','Days to expiry','In scope']
s=grid('Timesheet','Roster coverage','WORKLOAD    /    One row per master employee • selected month, current master attributes. Submission is matched by IQAMA, even if the submitted site differs.',heads,[34,18,24,26,17,17,18,20,16,16,14,18,18,18,18,14],E,(),'C6',7,'Timesheet_Input')
T='Timesheet_Input';P='Salarysheet';M='Employees';W='Workload Status';R='Timesheet';
for r in range(6,E+1):
 mm=lambda col:f"'Employees'!{col}{r}";crit=f'{rng(T,"A")},ViewMonth,{rng(T,"BM")},B{r}'
 cols={1:mm('B'),2:mm('Z'),3:mm('G'),4:mm('H'),5:mm('J'),6:mm('Q'),7:f'COUNTIFS({crit})',8:f'IF(G{r}=0,"Not submitted",IF(G{r}>1,"Duplicate",IF(SUMIFS({rng(P,"N")},{crit})>0,"Ready",IF(COUNTIFS({crit},{rng(P,"O")},"Ready")>0,"Ready","On hold"))))',9:f'SUMIFS({rng(T,"AM")},{crit})',10:f'SUMIFS({rng(T,"AO")},{crit})',11:f'SUMIFS({rng(T,"AN")},{crit})',12:f'SUMIFS({rng(P,"N")},{crit})',13:f'SUMIFS({rng(P,"M")},{crit})',14:mm('V'),15:f'IF(ISNUMBER({mm("E")}),{mm("E")},"")',16:mm('Y')}
 for j,f in cols.items():value(s,r,j,f'=IF({mm("B")}="","",{f})',fmt=MONEY if j in (12,13) else NUM if j in (7,9,10,11,15) else None)
flag(s,f'H6:H{E}','H6="On hold"');flag(s,f'H6:H{E}','H6="Ready"',MINT,GREEN);s['D3']='Reporting month';s['E3']='=ViewMonth';s['E3'].number_format=PER;hide(s,'P','P')
# Interactive consolidated client statement: period slicers + optional release-state selector.
heads=['Period','Client / site','Employee name','IQAMA','Release']+[f'{j:02}' for j in range(1,32)]+['Present eq.','Off days','Absent days','Logged days','Base','Overtime','Deductions','Net estimate','Held net','Ready net','Review item','Row index']
s=grid('Timesheet and Salarysheet','Client payroll statement','FILTERED STATEMENT    /    Follows Dashboard scope. Choose Ready or Hold below to print a release or review list. Daily marks link back to Attendance entry.',heads,[16,23,36,18,14]+[8]*31+[16]*10+[25,12],N,(),'F6',9,'Timesheet_Input');s['C3']='Release state  →';value(s,3,4,'All',True);dv(s,'D3','list',arr(['All','Ready','Hold']));s['E3']='Ready net →';merge(s,3,6,8,'=SUM(AT6:AT425)',GREEN,MINT,12,True).number_format=MONEY;s['I3']='Rows →';s['J3']='=COUNT(AV6:AV425)'
for r in range(6,N+1):
 value(s,r,48,f'=IFERROR(MATCH(ROW()-5,{rng(P,"W")},0),"")')
 for j in range(1,48):
  sh=P if j<6 or j>=37 else T;col= ['A','B','C','D','O'][j-1] if j<=5 else L(j+2) if j<=36 else ['R','T','S','Q','F','I','K','L','M','N','P'][j-37];ref=f'INDEX({rng(sh,col)},$AV{r})';f=f'IF(ISBLANK({ref}),"",{ref})' if 6<=j<=36 else ref
  value(s,r,j,f'=IF($AV{r}="","",{f})',fmt=PER if j==1 else MONEY if 41<=j<=46 else NUM if 37<=j<=40 else None,h='center' if 6<=j<=36 else None)
hide(s,'AV','AV');s.auto_filter.ref=f'A5:AU{N}';flag(s,f'E6:E{N}','E6="Hold"');flag(s,f'E6:E{N}','E6="Ready"',MINT,GREEN);s.print_area='A1:AU78'
# Expenses preserve source rows, including dated blanks, but only positive numeric claims enter metrics.
heads=['Claim date','Employee name','Client / worksite','Category','Amount • SAR','Claim status','Payment method','Description','Approval review','Period','Counted']
s=grid('Expense Manage','Expense register','CLAIMS    /    Enter an amount to activate a claim. Missing employee assignments remain unassigned; claim costs are not subtracted from payroll.',heads,[19,31,24,23,20,18,20,46,22,17,12],X,set(range(1,9)),'C6',6,'Control Center')
for r in range(6,X+1):
 row=r-2;vs=[ex.cell(row,j).value for j in [2,3,4,5,6,7,8,9]] if row<=123 else [None]*8
 for j,v in enumerate(vs,1):value(s,r,j,v,True,DATE if j==1 else MONEY if j==5 else None)
 for j,f in {9:f'IF(E{r}="","",IF(OR(NOT(ISNUMBER(E{r})),E{r}<=0,NOT(ISNUMBER(A{r}))),"Check entry",IF(E{r}>\'Control Center\'!$B$9,"HR review","Standard")))',10:f'IF(ISNUMBER(A{r}),DATE(YEAR(A{r}),MONTH(A{r}),1),"")',11:f'--AND(ISNUMBER(E{r}),E{r}>0,ISNUMBER(A{r}))'}.items():value(s,r,j,'='+f,fmt=PER if j==10 else None)
for col,opts in [('D',[raw['Control Center'].cell(r,20).value for r in range(116,124)]),('F',['Approved','Pending','Rejected']),('G',['Bank Transfer','Company Card','Cash'])]:dv(s,f'{col}6:{col}{X}','list',arr([i for i in opts if i]))
dv(s,f'E6:E{X}','decimal',.01,1e8,'between','Enter a positive numeric claim amount.');hide(s,'K','K');flag(s,f'I6:I{X}','I6="HR review"',SAND,GOLD);s['D3']='Month claims  →';s['E3']=f'=SUMIFS(E6:E{X},J6:J{X},ViewMonth,K6:K{X},1)';s['E3'].number_format=MONEY;s['F3']='Month count  →';s['G3']=f'=COUNTIFS(J6:J{X},ViewMonth,K6:K{X},1)'
# Site workload, coverage and settlement agree with the same scoped records used in Dashboard.
heads=['Worksite','Roster','Active','Submitted staff¹','Not submitted¹','Coverage¹','Payroll lines²','Ready lines','Hold lines','Ready net • SAR','Held net • SAR','Present eq.','Absent days','Attendance','Order']
s=grid(W,'Worksite scorecard','SITE ANALYSIS    /    ¹ By current master site and ID. ² By submitted site. Employee moves and unknown IDs can make these counts differ.',heads,[25,14,14,20,20,16,18,16,16,21,21,17,16,16,13],C,(),'B6',7,'Timesheet and Salarysheet')
for r in range(6,C+1):
 scope=f'{rng(P,"B")},$A{r},{rng(P,"V")},1';roster=f'{rng(R,"C",E)},$A{r},{rng(R,"P",E)},1';sp=lambda col:f'SUMIFS({rng(P,col)},{scope})'
 fs={1:f'IF(OR(\'Clients\'!A{r}="",AND(ViewSite<>"All",ViewSite<>\'Clients\'!A{r})),"",\'Clients\'!A{r})',2:f'COUNTIFS({roster})',3:f'COUNTIFS({roster},{rng(R,"E",E)},"Active")',4:f'COUNTIFS({roster},{rng(R,"G",E)},">0")',5:f'B{r}-D{r}',6:f'IFERROR(D{r}/B{r},"")',7:f'COUNTIFS({scope})',8:f'COUNTIFS({scope},{rng(P,"O")},"Ready")',9:f'COUNTIFS({scope},{rng(P,"O")},"Hold")',10:sp('N'),11:sp('M'),12:sp('R'),13:sp('S'),14:f'IFERROR(L{r}/(L{r}+M{r}),"")',15:f'IF(G{r}=0,"",1+COUNTIFS($G$6:$G${C},">0",$J$6:$J${C},">"&J{r})+COUNTIFS($G$6:G{r},">0",$J$6:J{r},J{r})-1)'}
 for j,f in fs.items():value(s,r,j,'='+f if j==1 else f'=IF($A{r}="","",{f})',fmt=PCT if j in (6,14) else MONEY if j in (10,11) else NUM if j!=1 else None)
cfbar(s,f'F6:F{C}',AQUA);cfbar(s,f'J6:J{C}','8ECABB');hide(s,'O','O');s['D3']='Reporting month';s['E3']='=ViewMonth';s['E3'].number_format=PER
# Audit: actionable filtered pay holds (one row per hold) and source employment-exception register.
heads=['Input row','Period','Submitted employee','IQAMA','Submitted site','First review item','Expected name','Expected site','Held net • SAR','', 'Master employee','Current site','Work status','Master IQAMA','Remarks','Hold index','Master index']
s=grid('Audit','Review queue','EXCEPTIONS    /    Left: first blocking reason per held payroll line in the selected scope. Right: current employment exceptions, independent of payroll submissions.',heads,[15,17,32,18,24,25,33,24,19,3,33,23,17,18,24,12,12],N,(),'D6',7,'Timesheet_Input');s['A4']='PAYROLL HOLDS';merge(s,4,2,9,'Resolve the identity, attendance or rate issue at the source. No entry is authorized by this view.',MUT,PALE);merge(s,4,11,15,'EMPLOYMENT EXCEPTIONS   /   Runaway, left, resigned, terminated or not interested.',WHITE,TEAL,10,True)
for r in range(6,N+1):
 value(s,r,16,f'=IFERROR(MATCH(ROW()-5,{rng(P,"X")},0),"")');value(s,r,17,f'=IFERROR(MATCH(ROW()-5,{rng(M,"AB",E)},0),"")')
 for j,col in [(1,''),(2,'A'),(3,'C'),(4,'D'),(5,'B'),(6,'P'),(7,'F'),(8,'G'),(9,'M'),(11,'B'),(12,'G'),(13,'J'),(14,'Z'),(15,'U')]:
  ix=f'$P{r}' if j<=9 else f'$Q{r}';sh=T if j in (7,8) else M if j>=11 else P;f=f'{ix}+5' if j==1 else f'INDEX({rng(sh,col,E if sh==M else N)},{ix})';value(s,r,j,f'=IF({ix}="","",{f})',fmt=MONEY if j==9 else PER if j==2 else None)
hide(s,'P','Q');s.auto_filter.ref=f'A5:I{N}';flag(s,f'F6:F{N}','F6<>""',SAND,GOLD)
# Interactive employee search with exact ID or exact name, independent of the roster filters.
s=wb[M];s['C3']='Find ID / name →';style(s['C3'],MUT,PALE,10,True);value(s,3,4,None,True,'@');merge(s,3,5,9,f'=IF(D3="","Enter an IQAMA or exact name in the blue field",IFERROR(INDEX($B$6:$B${E},MATCH(TRIM(D3&""),$Z$6:$Z${E},0))&"  /  "&INDEX($G$6:$G${E},MATCH(TRIM(D3&""),$Z$6:$Z${E},0)),IFERROR(INDEX($Z$6:$Z${E},MATCH(TRIM(D3),$B$6:$B${E},0))&"  /  "&INDEX($J$6:$J${E},MATCH(TRIM(D3),$B$6:$B${E},0)),"No exact match")))',TEAL,MINT,10)
# Slicer data: live bridges / chart sources / ranked records. No cached constants drive output.
for col,v in zip('JKLMN',['Site rank','Submitted site','Ready net','Held net','Lines']):ls[f'{col}1']=v
for r in range(2,8):
 ls[f'J{r}']=r-1
 for col,ref in zip('KLMN','AJKG'):ls[f'{col}{r}']=f'=IFERROR(INDEX({rng(W,ref,C)},MATCH(J{r},{rng(W,"O",C)},0)),"" )'
ls['P1']='Release state';ls['Q1']='Lines';ls['P2']='Ready';ls['P3']='On hold';ls['Q2']=f'=COUNTIFS({rng(P,"V")},1,{rng(P,"O")},"Ready")';ls['Q3']=f'=COUNTIFS({rng(P,"V")},1,{rng(P,"O")},"Hold")'
# Macro-free dashboard. Four aligned columns with short descriptive metric labels and visible scope.
d=wb['Dashboard'];setup(d);d.sheet_view.zoomScale=90;d.sheet_properties.tabColor=TEAL;d.sheet_format.defaultRowHeight=17
for j in range(1,18):d.column_dimensions[L(j)].width=2.5 if j in [1,5,9,13,17] else 10.5
for row in d.iter_rows(min_row=1,max_row=91,min_col=1,max_col=17):
 for c in row:c.fill=fill(BG)
for rr in (1, 5, 7, 10, 13, 18, 23, 86, 104, 111, 120, 24, 37):d.row_dimensions[rr].height=8
merge(d,2,2,16,'B A S M A     /     W O R K F O R C E    O P E R A T I O N S',AQUA,NAVY,10,True);d.row_dimensions[2].height=22
merge(d,3,2,16,'Workforce & payroll',WHITE,NAVY,26,True);d.row_dimensions[3].height=41
merge(d,4,2,16,'="MONTHLY CONTROL DESK     /     "&UPPER(TEXT(ViewMonth,"mmmm yyyy"))&"     •     "&IF(ViewSite="All","ALL WORKSITES",ViewSite)&"     •     SAR"',WHITE,NAVY,10);d.row_dimensions[4].height=24
starts=[2,6,10,14]
for c,sh,txt in zip(starts,['Timesheet_Input',P,'Audit','Control Center'],['01   ATTENDANCE ENTRY   →','02   PAYROLL LEDGER   →','03   REVIEW QUEUE   →','04   SETTINGS & GUIDE   →']):
 merge(d,6,c,c+2,txt);nav(d,L(c)+'6',sh,txt)
d.row_dimensions[6].height=22
for row,labels,initial,lists in [(8,['REPORTING MONTH','SUBMITTED WORKSITE','MASTER SPONSOR','MASTER WORK STATUS'],[start,'All','All','All'],['Periods','Sites','Sponsors','Statuses']),(11,['NATIONALITY','CONTRACT TYPE','RANK RECORDS BY','COMPLIANCE AS OF'],['All','All','Ready net',"='Control Center'!$B$10"],['Nations','Contracts','Sort',None])]:
 for c,lab,v,lis in zip(starts,labels,initial,lists):
  merge(d,row,c,c+2,lab,MUT,BG,9,True);cell=merge(d,row+1,c,c+2,v,BLUE if lis else MUT,SKY if lis else PALE,12,True);cell.border=Border(bottom=Side(style='thin',color='95BCBB'))
  if lis:dv(d,cell.coordinate,'list',lis,msg='Choose a report filter. Select All to reset this dimension.',blank=False)
  if lis=='Periods':cell.number_format='mmmm yyyy'
  if not lis:cell.number_format=DATE
 d.row_dimensions[row].height=16;d.row_dimensions[row+1].height=25

def sump(c):return f'SUMIF({rng(P,"V")},1,{rng(P,c)})'
def counte(col,val):return f'COUNTIFS({rng(M,"Y",E)},1,{rng(M,col,E)},"{val}")'
def card(row,i,label,formula,desc,fmt,theme='normal'):
 c=starts[i];bg,fg=((TEAL,WHITE) if theme=='hero' else (SAND,GOLD) if theme=='warning' else (WHITE,INK));merge(d,row,c,c+2,label,fg if theme!='normal' else MUT,bg,9,True);v=merge(d,row+1,c,c+2,'='+formula,fg,bg,23,True);v.number_format=fmt;merge(d,row+2,c,c+2,desc,fg if theme!='normal' else MUT,bg,10);d.row_dimensions[row].height=20;d.row_dimensions[row+1].height=35;d.row_dimensions[row+2].height=20
card(14,0,'READY NET   /   SAR',sump('N'),'Passed all release checks',MONEY,'hero');card(14,1,'HELD NET   /   SAR',sump('M'),'Separate from deductions',MONEY,'warning');card(14,2,'GROSS PAY   /   SAR',sump('J'),'Base pay plus overtime',MONEY);card(14,3,'ATTENDANCE RATE',f'IFERROR({sump("R")}/({sump("R")}+{sump("S")}),"—")','Present ÷ (present + absent)',PCT)
card(19,0,'PAYROLL RECORDS',f'SUM({rng(P,"V")})','Submitted lines in scope',INT);card(19,1,'RECORDS ON HOLD',"'Slicer Data'!Q3",'Open Review queue for reasons',INT,'warning');card(19,2,'TOTAL WORKFORCE',f'SUM({rng(M,"Y",E)})','Master employees in scope',INT);card(19,3,'SUBMISSION COVERAGE',f'IFERROR(COUNTIFS({rng(R,"P",E)},1,{rng(R,"G",E)},">0")/J20,"—")','Submitted IDs ÷ current roster',PCT)
# Two chart panels, anchored to independently sized row blocks.
band(d,82,'05    /    Payroll distribution',2,16);merge(d,83,2,8,'Net pay by submitted worksite',INK,WHITE,15,True);merge(d,83,10,16,'Release readiness',INK,WHITE,15,True);d.row_dimensions[83].height=27
merge(d,84,2,8,'TOP SIX   /   READY + HELD NET IN SAR',MUT,WHITE,9,True);merge(d,84,10,16,'RECORDS   /   PASSED CHECKS VS. ON HOLD',MUT,WHITE,9,True);d.row_dimensions[84].height=17
# Absolute pixel dimensions fit B:H and J:P at default workbook font metrics.
bc=BarChart();bc.type='bar';bc.grouping='stacked';bc.overlap=100;bc.gapWidth=45;bc.add_data(Reference(ls,min_col=12,max_col=13,min_row=1,max_row=7),titles_from_data=True);bc.set_categories(Reference(ls,min_col=11,min_row=2,max_row=7));bc.legend.position='b';bc.x_axis.title=None;bc.y_axis.title=None;bc.x_axis.scaling.min=0;bc.x_axis.numFmt='#,##0';bc.height=6.6;bc.width=12.1;bc.visible_cells_only=False;bc.y_axis.scaling.orientation='minMax';bc.x_axis.scaling.orientation='maxMin';bc.x_axis.crosses='max';bc.graphical_properties=GraphicalProperties(solidFill=WHITE,ln=__import__('openpyxl').drawing.line.LineProperties(noFill=True));
for se,co in zip(bc.series,[TEAL,'DDAD6B']):se.graphicalProperties.solidFill=co;se.graphicalProperties.line.noFill=True
bc.display_blanks='gap';bc.anchor=TwoCellAnchor(_from=AnchorMarker(col=1,row=84),to=AnchorMarker(col=8,row=98));d.add_chart(bc)
dc=DoughnutChart();dc.add_data(Reference(ls,min_col=17,min_row=1,max_row=3),titles_from_data=True);dc.set_categories(Reference(ls,min_col=16,min_row=2,max_row=3));dc.holeSize=72;dc.firstSliceAng=270;dc.width=12.1;dc.height=6.6;dc.legend.position='b';dc.visible_cells_only=False;dc.graphical_properties=GraphicalProperties(solidFill=WHITE);dc.series[0].data_points=[DataPoint(idx=i,spPr=GraphicalProperties(solidFill=co)) for i,co in enumerate([TEAL,'DDAD6B'])];dc.dataLabels=DataLabelList();dc.dataLabels.showVal=True;dc.dataLabels.showLegendKey=False;dc.dataLabels.showSerName=False;dc.dataLabels.showCatName=False;dc.dataLabels.showPercent=False;dc.anchor=TwoCellAnchor(_from=AnchorMarker(col=9,row=84),to=AnchorMarker(col=16,row=98));d.add_chart(dc)
for rr in range(85,99):d.row_dimensions[rr].height=14
# Scope and accounting checks are mathematically independent of chart display.
merge(d,99,2,16,'=IF(B20=0,"No payroll records in this scope. Change the month or filters; the master roster is shown independently.",IF(ABS(B15+F15-('+sump('L')+'))<0.005,"RECONCILED   /   Ready net + held net equals the ledger net estimate.","CHECK TOTALS   /   Net bridge does not reconcile."))',GREEN,MINT,11,True);d.row_dimensions[99].height=26
band(d,101,'06    /    Settlement detail',2,16)
for i,(lab,col,txt) in enumerate([('BASE PAY   /   SAR','F','Eligible base or estimate'),('OVERTIME   /   SAR','I','Cycle, approved hours or override'),('DEDUCTIONS   /   SAR','K','Capped at individual base pay'),('NET ESTIMATE   /   SAR','L','Before release controls')]):card(102,i,lab,sump(col),txt,MONEY)
band(d,106,'07    /    Workforce & compliance signals',2,16)
for i,(lab,form,txt,fmt) in enumerate([('ACTIVE STAFF',counte('J','Active'),'Current master work status',INT),('ON LEAVE',counte('J','On Leave'),'Current master work status',INT),('EXPIRED IQAMAS',counte('V','Expired'),'At the compliance date above',INT),('IQAMAS DUE SOON',counte('V','Due soon'),'Inside configured alert window',INT)]):card(107,i,lab,form,txt,fmt)
for i,(lab,form,txt,fmt) in enumerate([('PRESENT EQUIVALENTS',sump('R'),'0.5P counts as half a day',NUM),('RECORDED ABSENCES',sump('S'),'A + explicit numeric zero',NUM),('RECORDED OFF DAYS',sump('T'),'Excluded from attendance rate',INT),('READY RECORD SHARE',f'IFERROR(\'Slicer Data\'!Q2/B20,"—")','Ready lines ÷ submitted lines',PCT)]):card(112,i,lab,form,txt,fmt)
for i,(lab,form,txt,fmt) in enumerate([('INSURANCE EXPIRED',counte('W','Expired'),'Current roster, not pay holds',INT),('INSURANCE DUE SOON',counte('W','Due soon'),'Inside insurance alert window',INT),('FULL-TIME STAFF',counte('Q','Full Time'),'Current contract mix',INT),('AVG. READY PAY   /   SAR',f'IFERROR(B15/\'Slicer Data\'!Q2,"—")','Ready net ÷ ready records',MONEY)]):card(116,i,lab,form,txt,fmt)
band(d,121,'08    /    Top ten employees · ranked payroll records',2,16);merge(d,122,2,16,'="TOP TEN   /   "&UPPER(ViewSort)&"   •   Ties retain input order. Rankings follow all selected filters."',MUT,PALE,10);d.row_dimensions[122].height=22
segments=[(2,2,'No.'),(3,6,'Employee'),(7,9,'Worksite'),(10,11,'Release'),(12,13,'Absences'),(14,16,'=ViewSort&IF(OR(ViewSort="Net estimate",ViewSort="Ready net")," • SAR","")')]
for a,b,label in segments:merge(d,123,a,b,label,WHITE,NAVY,10,True)
for r in range(124,134):
 ix=f'$R{r}';d[f'R{r}']=f'=IFERROR(MATCH(ROW()-123,{rng(P,"Z")},0),"")'
 for (a,b,label),col in zip(segments,[None,'C','B','O','S','L']):
  f='ROW()-123' if col is None else f'IF(ViewSort="Attendance rate",TEXT(INDEX({rng(P,chr(89))},{ix}),"0.0%"),IF(ViewSort="Absence days",TEXT(INDEX({rng(P,chr(89))},{ix}),"0.##"),INDEX({rng(P,chr(89))},{ix})))' if col=='L' else f'INDEX({rng(P,col)},{ix})';c=merge(d,r,a,b,f'=IF({ix}="","",{f})',INK,WHITE if r%2==0 else PALE,10,col=='L',h='right' if col in ['L','S'] else 'left');c.number_format=MONEY if col=='L' else NUM if col=='S' else 'General'
 d.row_dimensions[r].height=22
hide(d,'R','R');band(d,24,'01    /    Submission & data quality',2,16)
for i,(lab,form,txt) in enumerate([('UNSUBMITTED STAFF',f'COUNTIFS({rng(R,"P",E)},1,{rng(R,"G",E)},0)','Master IDs without a period line'),('SUBMITTED WORKSITES',f'COUNTIF({rng(W,"G",C)},">0")','Configured or listed sites with lines'),('IDENTITY EXCEPTIONS',f'COUNTIFS({rng(P,"V")},1,{rng(P,"E")},"<>Matched")','Includes missing and duplicate IDs'),('SALARIES TO REVIEW',counte('X','Check salary'),'Missing, text or zero master pay')]):card(25,i,lab,form,txt,INT)
# Fixed expiry windows complement, rather than replace, the configurable due-soon alert.
# Freeze the interpretation of short windows in the visible labels: both endpoints inclusive.
band(d,30,'02    /    Workforce totals & Iqama expiry',2,16)
link=merge(d,119,2,16,'Open workforce totals & Iqama expiry windows  →',TEAL,MINT,10,True)
link.hyperlink=Hyperlink(ref='B119',location="'Dashboard'!B30",display=link.value);d.row_dimensions[119].height=21
card(31,0,'TOTAL INACTIVE',counte('J','Suspended')+'+'+counte('J','Terminated'),'Suspended + terminated; no leave',INT)
card(31,1,'TOTAL RANKED RECORDS',f'COUNT({rng(P,"Z")})','All ranked lines, not just top ten',INT)
def expiry_window(lo,hi):
 return f'COUNTIFS({rng(M,"Y",E)},1,{rng(M,"V",E)},"<>Missing date",{rng(M,"E",E)},">={lo}",{rng(M,"E",E)},"<={hi}")'
windows=[(31,2,9,15,'IQAMA  /  9–15 DAYS'),(31,3,1,14,'IQAMA  /  1–14 DAYS')]
for row,limits in [(37,[60,30,20,15]),(41,[14,9,1,0])]:
 for i,hi in enumerate(limits):windows.append((row,i,0,hi,'IQAMA  /  DUE TODAY' if hi==0 else f'IQAMA  /  WITHIN {hi} DAY'+('S' if hi!=1 else '')))
merge(d,35,2,16,'="EXPIRY WINDOWS   /   As of "&TEXT(N12,"dd mmm yyyy")&". Counts overlap: do not add. Expired and missing dates are excluded."',MUT,PALE,10,True);d.row_dimensions[35].height=23
for row,i,lo,hi,label in windows:
 caption=('Expires on the reference date' if hi==0 else f'{lo}–{hi} days remaining, inclusive')
 card(row,i,label,expiry_window(lo,hi),caption,'#,##0','warning' if hi<=1 else 'normal')
for rr in (29, 34, 36, 40, 44):d.row_dimensions[rr].height=8
# A shared text-color key. Apply differential font colors only, preserving numeric
# formats, KPI type size, blue input backgrounds and the existing protection rules.
OK_TEXT='176244'; WARN_TEXT='856000'; BAD_TEXT='A62C35'; YELLOW='FFF3C4'
def ink_rule(sh,area,test,color):
 sh.conditional_formatting.add(area,FormulaRule(formula=[test],font=Font(color=color)))
for key,ref in [('IqamaWarnDays',7),('InsuranceWarnDays',8),('ReviewDate',10)]:nm(key,f"'Control Center'!$B${ref}")
for sh in (master,wb[M]):
 for area in (f'D6:D{E}',f'E6:E{E}'):
  for test,col in [('$E6<0',BAD_TEXT),('AND($E6>=0,$E6<=IqamaWarnDays)',WARN_TEXT),('$E6>IqamaWarnDays',OK_TEXT)]:
   ink_rule(sh,area,f'AND($B6<>"",ISNUMBER($E6),{test})',col)
 for test,col in [('INT($R6)<INT(ReviewDate)',BAD_TEXT),('AND(INT($R6)>=INT(ReviewDate),INT($R6)<=INT(ReviewDate)+InsuranceWarnDays)',WARN_TEXT),('INT($R6)>INT(ReviewDate)+InsuranceWarnDays',OK_TEXT)]:
  ink_rule(sh,f'R6:R{E}',f'AND($B6<>"",ISNUMBER($R6),$R6>0,ISNUMBER(ReviewDate),ReviewDate>0,{test})',col)
 for status,col in [('Active',OK_TEXT),('On Leave',WARN_TEXT),('Suspended',WARN_TEXT),('Terminated',BAD_TEXT)]:ink_rule(sh,f'J6:J{E}',f'$J6="{status}"',col)
flag(wb[M],f'V6:W{E}','V6="Current"',MINT,OK_TEXT)
ink_rule(wb[M],f'V6:W{E}','V6="Missing date"',MUT)
for stat,col in [('Expired',BAD_TEXT),('Due soon',WARN_TEXT),('Current',OK_TEXT),('Missing date',MUT)]:
 ink_rule(wb[R],f'N6:O{E}',f'$N6="{stat}"',col)
for stat,col in [('Ready',OK_TEXT),('Hold',BAD_TEXT)]:ink_rule(d,'J124:J133',f'$J124="{stat}"',col)
for sh,area,choice,co in [(T,f'E6:E{N}','Matched',OK_TEXT),(T,f'AQ6:AQ{N}','Complete',OK_TEXT),('Clients',f'O6:O{C}','Configured',OK_TEXT)]:
 ink_rule(wb[sh],area,area.split(':')[0]+'="'+choice+'"',co)
for addr,co in [('J108',BAD_TEXT),('B117',BAD_TEXT),('N108',WARN_TEXT),('F117',WARN_TEXT),('F20',BAD_TEXT),('J26',BAD_TEXT),('N26',WARN_TEXT),('B108',OK_TEXT)]+[(L(starts[i])+str(row+1),WARN_TEXT) for row,i,lo,hi,_ in windows]:
 ink_rule(d,addr,f'AND(ISNUMBER({addr}),{addr}>0)',co)
# Empty space above the payroll charts becomes a compact key and a jump link.
for c,txt,fg,bg in [(2,'RED   /   Expired or blocked',BAD_TEXT,ROSE),(6,'YELLOW   /   Due or in review',WARN_TEXT,YELLOW),(10,'GREEN   /   Current or ready',OK_TEXT,MINT)]:merge(d,23,c,c+2,txt,fg,bg,9,True)
link=merge(d,23,14,16,'Workforce insights  →',TEAL,PALE,10,True);link.hyperlink=Hyperlink(ref='N23',location="'Dashboard'!B45",display=link.value);d.row_dimensions[23].height=23
# Non-overlapping categories: these charts reconcile to the in-scope staff count,
# unlike the intentionally overlapping renewal KPI windows above. Both use the
# linked employee directory; they never infer workforce composition from payroll.
ls['S1']='Iqama expiry';ls['T1']='Employees';ls['V1']='Work status';ls['W1']='Employees'
age_bands=[('Expired',counte('V','Expired')),('0–15 days',expiry_window(0,15)),('16–30 days',expiry_window(16,30)),('31–60 days',expiry_window(31,60)),('Over 60 days',f'COUNTIFS({rng(M,"Y",E)},1,{rng(M,"E",E)},">60")'),('Missing date',counte('V','Missing date'))]
for rr,(label,formula) in enumerate(age_bands,2):ls[f'S{rr}']=label;ls[f'T{rr}']='='+formula
for rr,label in enumerate(['Active','On Leave','Suspended','Terminated','Unclassified'],2):
 ls[f'V{rr}']=label;ls[f'W{rr}']='='+counte('J',label) if label!='Unclassified' else "='Dashboard'!J20-SUM(W2:W5)"
band(d,45,'03    /    Workforce insights',2,16)
merge(d,46,2,8,'Iqama renewal horizon',INK,WHITE,15,True);merge(d,46,10,16,'Workforce by status',INK,WHITE,15,True);d.row_dimensions[46].height=27
merge(d,47,2,8,'DISTINCT EXPIRY BANDS   /   EMPLOYEES',MUT,WHITE,9,True);merge(d,47,10,16,'CURRENT ROSTER   /   EMPLOYEES',MUT,WHITE,9,True);d.row_dimensions[47].height=18

def profile_chart(c1,c2,cat_col,num_col,last,colors):
 ch=BarChart();ch.type='bar';ch.grouping='clustered';ch.gapWidth=58
 ch.add_data(Reference(ls,min_col=num_col,min_row=1,max_row=last),titles_from_data=True);ch.set_categories(Reference(ls,min_col=cat_col,min_row=2,max_row=last))
 ch.legend=None;ch.visible_cells_only=False;ch.display_blanks='gap';ch.x_axis.scaling.orientation='maxMin';ch.y_axis.scaling.min=0;ch.y_axis.crosses='max';ch.y_axis.numFmt='0';ch.x_axis.tickLblPos='low';ch.y_axis.majorGridlines=None
 ch.graphical_properties=GraphicalProperties(solidFill=WHITE);series=ch.series[0];series.graphicalProperties.line.noFill=True
 series.data_points=[DataPoint(idx=i,spPr=GraphicalProperties(solidFill=co,ln=__import__('openpyxl').drawing.line.LineProperties(noFill=True))) for i,co in enumerate(colors)]
 ch.dataLabels=DataLabelList();ch.dataLabels.dLblPos='outEnd';ch.dataLabels.showVal=True;ch.dataLabels.showCatName=False;ch.dataLabels.showSerName=False;ch.dataLabels.showLegendKey=False;ch.dataLabels.showPercent=False;ch.dataLabels.numFmt='0'
 ch.anchor=TwoCellAnchor(_from=AnchorMarker(col=c1,row=47),to=AnchorMarker(col=c2,row=61));d.add_chart(ch);return ch
aging_chart=profile_chart(1,8,19,20,7,['BA4952','D5A427','E3C267','8FA6A5','3D9072','BBC7C7'])
status_chart=profile_chart(9,16,22,23,6,['3D9072','E3C267','D5A427','BA4952','BBC7C7'])
for rr in range(48,62):d.row_dimensions[rr].height=14
for c,txt in [(2,'=IF(J20=0,"No employees match the selected roster filters.",TEXT(J108/J20,"0.0%")&" of the roster has an expired Iqama. "&TEXT(\'Slicer Data\'!T3,"#,##0")&" expire in the next 0–15 days.")'),(10,'=IF(J20=0,"No employees match the selected roster filters.",TEXT(B108,"#,##0")&" of "&TEXT(J20,"#,##0")&" employees are active ("&TEXT(B108/J20,"0.0%")&"). "&TEXT(B32,"#,##0")&" are suspended or terminated.")')]:
 cell=merge(d,62,c,c+6,txt,INK,PALE,11,True);cell.alignment=align('left',1,True)
d.row_dimensions[62].height=38;d.row_dimensions[63].height=8
band(d,135,'09    /    Workbook navigation',2,16)
for r,links in [(136,[(W,'Worksite scorecard'),('Timesheet and Salarysheet','Filtered client statement'),(R,'Roster coverage'),(M,'Employee directory') ]),(138,[('Clients','Client rate cards'),('Expense Manage','Expense register'),('Audit','Payroll & employment audit'),('Control Center','Rules & source notes')])]:
 for j,(sh,label) in zip(starts,links):merge(d,r,j,j+2,label);nav(d,f'{L(j)}{r}',sh,label+'  →')
 d.row_dimensions[r].height=27
merge(d,141,2,16,'BASMAT AL-MAWARED   /   INTERNAL OPERATIONS     •     Blue cells are editable. Held pay is an estimate, not an authorized transfer.',MUT,PALE,10);d.row_dimensions[141].height=25;d.print_area='A1:Q123';d.page_setup.fitToHeight=2;d.sheet_view.selection[0].activeCell='B9';d.sheet_view.selection[0].sqref='B9'
# Year view: no activity is distinguished from a zero-valued posted month.
s=wb[W];band(s,48,'YEARBOOK    /    Monthly payroll and expense activity',1,10);merge(s,49,1,11,'Payroll follows the scope filters except month. Expense figures are company-wide for each month; unassigned claims cannot be attributed to a site.',MUT,PALE,10)
for j,txt in enumerate(['Month','Payroll lines','Ready lines','Hold lines','Gross • SAR','Deductions','Net estimate','Ready net','Held net','Claims • all sites','Expenses • all sites'],1):style(s.cell(50,j,txt),WHITE,NAVY,10,True)
for r in range(51,63):
 value(s,r,1,f'=DATE(YEAR(ViewMonth),{r-50},1)',fmt='mmm yyyy')
 crit=f'{rng(T,"A")},A{r},{rng(T,"BP")},1'
 for j,f in {2:f'COUNTIFS({crit})',3:f'COUNTIFS({crit},{rng(P,"O")},"Ready")',4:f'COUNTIFS({crit},{rng(P,"O")},"Hold")',**{k:f'IF(B{r}=0,"",SUMIFS({rng(P,cl)},{crit}))' for k,cl in zip(range(5,10),'JKLNM')},10:f'COUNTIFS({rng("Expense Manage","J",X)},A{r},{rng("Expense Manage","K",X)},1)',11:f'IF(J{r}=0,"",SUMIFS({rng("Expense Manage","E",X)},{rng("Expense Manage","J",X)},A{r},{rng("Expense Manage","K",X)},1))'}.items():value(s,r,j,'='+f,fmt=MONEY if j in (5,6,7,8,9,11) else INT)
s.print_area='A1:N63';wb[T]['BP5']='Dimension slice'
for r in range(6,N+1):wb[T][f'BP{r}']=wb[T][f'BE{r}'].value.replace(f'A{r}=ViewMonth,','')
hide(wb[T],'BP','BP')
# Dimension dropdowns follow the master rather than a fixed list of original sponsors.
for col,master,key,listcol in [('AC','H','Sponsors','B'),('AD','M','Nations','C'),('AE','Q','Contracts','D')]:
 em=wb[M];em[f'{col}5']=key+' list order'
 for r in range(6,E+1):em[f'{col}{r}']=f'=IF(OR(B{r}="",{master}{r}=""),"",IF(COUNTIF(${master}$6:{master}{r},{master}{r})=1,COUNT(${col}$5:{col}{r-1})+1,""))'
 ls[f'{listcol}3']='Unmatched'
 for r in range(4,354):ls[f'{listcol}{r}']=f'=IFERROR(INDEX({rng(M,master,E)},MATCH(ROW()-3,{rng(M,col,E)},0)),"")'
 nm(key,f"'Slicer Data'!${listcol}$2:INDEX('Slicer Data'!${listcol}$2:${listcol}$353,2+COUNT({rng(M,col,E)}))")
hide(wb[M],'AC','AE')
# Preserve the source's settings identity separately from the operational dashboard brand.
guide(55,'Source settings identity','The supplied Control Center identifies the company as Al Manar HR Solutions; the supplied operational dashboard is branded Basmat Al-Mawared. The operational brand is retained here.')
band(g,57,'05    /    Verification record')
guide(58,'Imported baseline • September','The September build was recalculated and checked independently for all 73 rows: 55 passed and 18 were held. Baseline ready SAR 75,630.06 + held SAR 23,163.68 = net estimate SAR 98,793.74. These are build checks, not live KPIs.')
guide(59,'Behavior tests','Test copies checked site filtering, ranked results, hold-only statements, empty scopes, new master values, leap February, duplicates, missing days, invalid codes, OT overrides, cycle premiums, half-pay absences and capped deductions.')
guide(60,'Closing and printing','Save a separate copy before changing master data or approving a month. Register print areas initially cover the imported data; extend the print area when additional records need to be printed.')
guide(61,'Workforce totals','Total workforce is the current roster in scope. Inactive means Suspended or Terminated; On Leave is separate. Total ranked records counts all in-scope payroll lines with a numeric rank, including those below the first ten.')
guide(62,'Iqama planning windows','Within N days means 0 through N calendar days from the compliance date. The 9–15 and 1–14 day windows include both endpoints and exclude today. The windows overlap, are not additive, and exclude past or unknown dates.')
g.print_area='A1:H62'
# The audit row locator opens the corresponding entry, even after sorting the exception view.
for r in range(6,N+1):wb['Audit'][f'A{r}']=f'=IF($P{r}="","",HYPERLINK("#\'Timesheet_Input\'!D"&($P{r}+5),"Open row "&($P{r}+5)))'
# Evidence and transfer upgrade: retain the existing time and pay engines. Each
# fact has one input location. The Main Dataset hub joins their records, it is
# not a second editable timesheet. No inferred corrections to a worker's data.
master=wb['Main Dataset'];audit=wb['Audit'];tim=wb[T];pay=wb[P];unlock=xl.styles.Protection(locked=False)
def jump(sh,addr,tab,dst,txt):
 ce=sh[addr];ce.value=txt;ce.hyperlink=Hyperlink(ref=addr,location=f"'{tab}'!{dst}",display=txt);style(ce,TEAL,MINT,10,True)
def linked(tab,addr):
 ref=f"'{tab}'!{addr}";return f'=IF(ISBLANK({ref}),"",{ref})'
def hdr(sh,r,j,label):style(sh.cell(r,j,label),WHITE,NAVY,10,True,'left',1,True)
def view_lock(sh):sh.protection.sheet=True;sh.protection.autoFilter=False
# Every suspect record has a specific repair instruction. Dates for reports are
# distinct from entry dates and actual review dates. No TODAY() as a timestamp.
band(g,64,'06    /    Date controls & reporting calendar')
for r,lab,f,fmt,desc in [(65,'Report starts','=ViewMonth',DATE,'Dashboard selection • a filter, not an update to any time entry.'),(66,'Report ends','=DATE(YEAR(B65),MONTH(B65)+1,0)',DATE,'Last valid date in this reporting month; leap years are calculated.'),(67,'Calendar days','=DAY(B66)',INT,'Worked shifts may differ from calendar days or contractual pay days.'),(68,'Compliance as of','=ReviewDate',DATE,'Independent date used for Iqama and insurance checks.'),(69,'Next period','=B66+1',DATE,'Use this month-start date for new input lines; keep prior periods intact.')]:
 value(g,r,1,lab);value(g,r,2,f,fmt=fmt);merge(g,r,3,8,desc,MUT,PALE);g.row_dimensions[r].height=23
band(g,71,'07    /    Correction workflow & rule directory')
guide(72,'Where to enter data','Use Main Dataset for staff, Timesheet_Input for attendance, Clients for contractual rates, Expense Manage for claims and Control Center for policy. All other copies are linked outputs. The Main Dataset index opens the combined register and the source for each fact.')
guide(73,'Remake a held entry','Open Audit. Select All periods to include undated entries. Open the original time row and correct only what source evidence supports; do not create a second payable line. Correct a master or client error in its own source. All 31 marks and adjustments flow into the same payroll row.')
guide(74,'Review and recheck','Open the fixed-row Rework log on Audit. Enter reviewer, actual review date and evidence; copy Current revision to Signed revision using Paste Values. This cannot bypass rule failures. A subsequent change to the signed facts invalidates the review and holds the line again.')
guide(75,'Preserve evidence','The log retains the imported time fields and initial finding for the 73 supplied lines. New lines have live checks; start a review note before correcting a new suspect line. Keep a dated copy for closed periods and successive review history. No automatic or immutable edit history is claimed.')
rule_actions={
 'ID missing':'Enter the IQAMA shown on the verified source document.', 'Invalid IQAMA':'Check all ten IQAMA digits against the document. Do not infer missing digits.',
 'ID not found':'Verify the ID against Main Dataset; add a real master record only after confirmation.', 'Duplicate entry':'Compare both entries for the same month and ID. Consolidate the valid time into one line; remove only a confirmed duplicate after preserving evidence.', 'Duplicate master ID':'Resolve the duplicate IQAMA in Main Dataset against original documents.',
 'Name mismatch':'Check the submitted name against the matched master. Correct the erroneous source, not both.', 'Site mismatch':'Confirm the worksite assignment with the supporting timesheet before updating it.', 'Name & site mismatch':'Confirm identity and assignment; the expected master name and site are in the review queue.',
 'Invalid period':'Use the first day of the actual month as a real date. Do not change Dashboard month to repair an entry.',
 'Invalid day code':'Replace unsupported entries with P, A, OFF, 0.5P or valid numeric units for the pay basis.', 'Outside month':'Remove the entry outside this calendar month or enter it on a separate line for its actual month.', 'Incomplete days':'Complete missing days using the approved attendance record; never count blanks as absences.',
 'Confirm 30 Sep':'Compare both 30 September marks. Set the approved day-30 mark and confirm the date review.',
 'Check adjustment':'Correct negative or nonnumeric OT / deduction inputs using supporting evidence.', 'OT not allowed':'Confirm the contractual OT policy; do not pay an override against a no-OT rule.', 'OT hours unused':'Use the cycle premium rule or confirm the contractual rate card; do not add unused hours.',
 'Check client rates':'Complete the client salary, shift, absence and overtime rules in Clients.', 'Check master salary':'Confirm and enter the numeric employee salary in Main Dataset.',
 'Document correction':'Enter reviewer, review date and evidence. Paste the current revision as a value into Signed revision.',
 'Checks passed':'All spreadsheet checks passed. Separate authorization is still required before payment.'}
for rr,(rule,act) in enumerate(rule_actions.items(),78):value(g,rr,1,rule);c=merge(g,rr,2,8,act,INK,PALE if rr%2 else WHITE);c.alignment=align('left',1,True);g.row_dimensions[rr].height=29
rule_end=77+len(rule_actions);g.print_area=f'A1:H{rule_end}';nm('RepairRules',f"'Control Center'!$A$78:$B${rule_end}")
# Audit mode includes undated/invalid periods without altering the Dashboard slice.
audit['C3']='Show payroll holds →';value(audit,3,4,'Selected scope',True);audit['D3'].protection=unlock;dv(audit,'D3','list',arr(['Selected scope','All periods']),blank=False)
jump(audit,'E3','Audit','R4','REWORK LOG  →');jump(audit,'F3','Control Center','A71','REPAIR RULES  →')
audit['A2']='PAYROLL HOLDS   /   Choose Selected scope or All periods. The fixed-row Rework log retains imported evidence and signed corrections independently of the queue.'
merge(audit,4,18,25,'REWORK LOG   /   Fixed time rows • correct at source, record evidence, then sign the checked revision.',WHITE,TEAL,11,True)
for j,(txt,wi) in enumerate([('Time row',17),('Rework state',22),('Rule finding',23),('How to resolve',48),('Reviewed by',22),('Reviewed on',19),('Evidence / correction',42),('Current revision · copy',30),('Signed revision · values only',30),('Imported finding',23),('Source changed',18),('Review needed',16),('Review valid',16)],18):
 hdr(audit,5,j,txt);audit.column_dimensions[L(j)].width=wi
for j in (22,23,24,26):style(audit.cell(5,j),WHITE,TEAL,10,True,'left',1,True)
# Before-image snapshots are evidence only: no live calculation reads a historic
# salary or time amount instead of the current entry. Source values stay typed.
rawcols=[1,2,3,4,*range(8,39),44,45,46,47,48,49,50];snap0=31
merge(audit,4,snap0,snap0+len(rawcols)-1,'RECEIVED EVIDENCE   /   Original imported time fields • read only • never a payroll input',WHITE,NAVY,11,True)
for k,j in enumerate(rawcols):hdr(audit,5,snap0+k,'Received '+str(tim.cell(5,j).value));audit.column_dimensions[L(snap0+k)].width=16 if j not in (2,3,47) else 30
masters={cleanid(v[2]):v for v in ep};baseline_count=0
for r in range(6,N+1):
 tx=lambda col:f"'{T}'!{col}{r}";px=lambda col:f"'{P}'!{col}{r}"
 if r<=78:
  old=masters.get(tim[f'D{r}'].value);nam=tim[f'C{r}'].value;site=tim[f'B{r}'].value
  norm=lambda v:' '.join(str(v).replace('\xa0',' ').upper().split());nmatch=old is not None and norm(old[1])==norm(nam);smatch=old is not None and norm(old[6])==norm(site)
  reason='ID not found' if old is None else 'Name & site mismatch' if not nmatch and not smatch else 'Name mismatch' if not nmatch else 'Site mismatch' if not smatch else 'Confirm 30 Sep' if tim[f'BO{r}'].value==1 else 'Checks passed'
  value(audit,r,27,reason);baseline_count+=reason!='Checks passed'
  for k,j in enumerate(rawcols):value(audit,r,snap0+k,tim.cell(r,j).value,fmt=PER if j==1 else '@' if j==4 else None)
 # Blank and explicit numeric zero produce different strings; numeric and text
 # day marks are distinguished for approvals. Revision includes effective rates
 # and master identity, in addition to every money-affecting time input.
 tracked=[tx(L(j)) for j in [1,2,3,4,*range(8,39),44,45,46,50,6,7,64]]+[px(co) for co in ['AA','AB','AC','AD','AE','AF','AG','F','I','K','AJ']]
 def term(ref):return f'IF(ISNUMBER({ref}),"n"&{ref},IF({ref}="","_","s"&{ref}))'
 sig='&"|"&'.join(term(ref) for ref in tracked)
 changed=','.join(f'NOT(EXACT({tx(L(j))}&"",{L(snap0+k)}{r}&""))' for k,j in enumerate(rawcols))
 fields={18:f'IF({tx("BN")}=0,"",HYPERLINK("#\'Timesheet_Input\'!D{r}","Time row {r}"))',19:f'IF({tx("BN")}=0,"",IF({px("AJ")}<>"Checks passed","Correct source",IF(AC{r}=0,"Clear",IF(AD{r}=1,"Remade","Review required"))))',20:f'IF({tx("BN")}=0,"",{px("AJ")})',21:f'IF({tx("BN")}=0,"",IFERROR(VLOOKUP({px("P")},RepairRules,2,FALSE),"Inspect the source inputs."))',25:f'IF({tx("BN")}=0,"","T{r}|"&{sig})',28:f'IF(OR(AA{r}="",{tx("BN")}=0),0,--OR({changed}))',29:f'IF({tx("BN")}=0,0,--OR(AND(AA{r}<>"",AA{r}<>"Checks passed"),AB{r}=1,COUNTA(V{r}:X{r},Z{r})>0))',30:f'IF({tx("BN")}=0,0,--IFERROR(AND(LEN(TRIM(V{r}))>0,ISNUMBER(W{r}),W{r}>0,W{r}<=TODAY(),W{r}=INT(W{r}),LEN(TRIM(X{r}))>0,LEN(Z{r})>0,NOT(ISFORMULA(Z{r})),EXACT(Z{r},Y{r})),FALSE))'}
 # OOXML stores ISFORMULA with its compatibility prefix. A live reference
 # to the current revision is never accepted as a signed value snapshot.
 fields[30]=fields[30].replace('ISFORMULA(', '_xlfn.ISFORMULA(')
 for j,f in fields.items():value(audit,r,j,'='+f,fmt=INT if j>=28 else None)
 for j in [22,23,24,26]:value(audit,r,j,None,True,DATE if j==23 else '@' if j==26 else None).protection=unlock
 # Display every original time field on the payroll line, not just its totals.
 # AJ is the technical result before audit evidence; AK groups the final hold.
 base=pay[f'P{r}'].value;value(pay,r,36,base)
 pay[f'P{r}']=f'=IF({tx("BN")}=0,"",IF(AJ{r}<>"Checks passed",AJ{r},IF(AND(\'Audit\'!AC{r}=1,\'Audit\'!AD{r}<>1),"Document correction","Checks passed")))'
 pay[f'X{r}']=f'=IF({tx("BN")}=0,"",IF(AND(OR(V{r}=1,\'Audit\'!$D$3="All periods"),O{r}="Hold"),COUNT($X$5:X{r-1})+1,""))'
 cat=f'IF(O{r}<>"Hold","",IF(P{r}="Document correction","Review evidence",IF(E{r}<>"Matched","Identity / site",IF(OR(P{r}="Check client rates",P{r}="Check master salary",P{r}="Check adjustment",P{r}="OT not allowed",P{r}="OT hours unused"),"Rates / adjustments","Time / calendar"))))';value(pay,r,37,'='+cat)
 trace_cols=[*range(8,39),44,45,46,47,48,49,50,6,7,43,58]
 for k,j in enumerate(trace_cols,38):value(pay,r,k,linked(T,f'{L(j)}{r}'),fmt=MONEY if j in [45,46] else None,h='center' if j in range(8,39) else None)
 # Add a visible source-side rework flag without making it an input to validation.
 for j,f in [(69,f'IF(BN{r}=0,"",{px("O")})'),(70,f'IF(BN{r}=0,"",{px("P")})'),(71,f'IF(BN{r}=0,"",HYPERLINK("#\'Audit\'!V{r}","Review line {r}"))')]:value(tim,r,j,'='+f)
 # Imported evidence remains addressable even if editable identity fields are cleared.
 tim[f'BN{r}']=f'=--(COUNTA(A{r}:D{r},H{r}:AL{r},AR{r}:AT{r},AW{r})>0)'
 old_check=tim[f'E{r}'].value[1:];tim[f'E{r}']=f'=IF(AND(BN{r}=1,D{r}<>"",NOT(IFERROR(AND(LEN(BM{r})=10,TEXT(VALUE(BM{r}),"0000000000")=BM{r}),FALSE))),"Invalid IQAMA",{old_check})'
assert baseline_count==18
for j,tit in enumerate(['Pre-review check','Hold category']+[str(tim.cell(5,j).value) for j in trace_cols],36):hdr(pay,5,j,tit);pay.column_dimensions[L(j)].width=10 if 38<=j<=68 else 23
merge(pay,4,38,68,'TRANSFERRED ATTENDANCE   /   Identical source row • day 01–31 • blanks remain blank',WHITE,TEAL,10,True)
merge(pay,4,69,37+len(trace_cols),'SOURCE EVIDENCE   /   OT, adjustments, imported notes, verification and calendar checks',WHITE,NAVY,10,True)
for j,txt in [(69,'Payroll release'),(70,'Payroll finding'),(71,'Rework record')]:hdr(tim,5,j,txt);tim.column_dimensions[L(j)].width=24
jump(pay,'D3',P,'AL5','DAILY ENTRIES  →');jump(pay,'E3','Main Dataset','A360','DATA HUB  →');jump(tim,'C3',T,'BQ5','PAYROLL CHECKS  →');jump(tim,'D3','Audit','R4','REWORK LOG  →')
flag(tim,f'BQ6:BQ{N}','BQ6="Hold"');flag(tim,f'BQ6:BQ{N}','BQ6="Ready"',MINT,GREEN);flag(audit,f'S6:S{N}','S6="Correct source"');flag(audit,f'S6:S{N}','S6="Review required"',SAND,GOLD);flag(audit,f'S6:S{N}','S6="Remade"',MINT,GREEN)
dv(audit,f'W6:W{N}','date','DATE(1900,1,1)','TODAY()','between','Enter the actual completed review date; future dates are not accepted.')
# Protect the linked ledger, received snapshots and stable row mapping. Filters,
# selectors and designated entry cells stay editable; there is no password.
for rr in range(6,N+1):
 for j in [1,2,3,4,*range(8,39),44,45,46,47,50]:tim.cell(rr,j).protection=unlock
for sh in [audit,tim,pay]:view_lock(sh)
# One browsable source hub with distinct grains. Lower sections are calculated
# references, so time, rate and expense data are never entered a second time.
master['A1']='Main dataset & source hub';master['A2']='EMPLOYEE MASTER   /   Edit staff here. Use the links below for time, pay, audit, rate cards and claims. Blue cells are inputs; linked register rows are read only.'
for addr,tab,dst,txt in [('C3','Main Dataset','A360','TIME + PAY  →'),('D3','Main Dataset','A790','RATE CARDS  →'),('E3','Main Dataset','A840','CLAIMS  →'),('F3','Audit','R4','REWORK LOG  →'),('G3','Control Center','A64','DATE RULES  →')]:jump(master,addr,tab,dst,txt)
band(master,360,'SOURCE HUB    /    Time → payroll → audit   •   all periods; one line per time row',1,21)
merge(master,361,1,21,'Enter days only in Timesheet_Input. Open the payroll row for every transferred mark. Correct a suspect line at its source; review evidence is attached to the fixed time row, not the moving audit queue.',MUT,PALE,10)
for j,h in enumerate(['Time row','Submitted employee','IQAMA','Period','Payroll link','Submitted site','Release','First finding','Review state','Logged days','Calendar days','Present eq.','Absent days','Off days','Base pay','Overtime','Deductions','Net estimate','Held net','Ready net','Audit link'],1):hdr(master,362,j,h)
for r in range(6,N+1):
 rr=r+357
 for j,sh,co in [(2,P,'C'),(3,P,'D'),(4,P,'A'),(6,P,'B'),(7,P,'O'),(8,P,'P'),(9,'Audit','S'),(10,P,'Q'),(11,T,'BF'),(12,P,'R'),(13,P,'S'),(14,P,'T'),(15,P,'F'),(16,P,'I'),(17,P,'K'),(18,P,'L'),(19,P,'M'),(20,P,'N')]:
  value(master,rr,j,f'=IF(\'{T}\'!BN{r}=0,"",\'{sh}\'!{co}{r})',fmt=PER if j==4 else MONEY if 15<=j<=20 else None)
 for j,sh,dst,label in [(1,T,'D','Time'),(5,P,'AL','Daily'),(21,'Audit','R','Audit')]:
  value(master,rr,j,f'=IF(\'{T}\'!BN{r}=0,"",HYPERLINK("#\'{sh}\'!{dst}{r}","{label} {r}"))')
flag(master,'G363:G782','G363="Hold"');flag(master,'G363:G782','G363="Ready"',MINT,GREEN)
for top,srcsh,lastcol,end,label in [(790,'Clients',15,C,'Client contracts & pay rules'),(840,'Expense Manage',11,X,'Expense claims & approval checks')]:
 band(master,top,'SOURCE HUB    /    '+label,1,lastcol);jump(master,'A'+str(top+1),srcsh,'A5','OPEN SOURCE  →');merge(master,top+1,2,lastcol,'Live view of the authoritative register. Edit at the source, not in this linked view.',MUT,PALE)
 for j in range(1,lastcol+1):hdr(master,top+2,j,wb[srcsh].cell(5,j).value)
 for r in range(6,end+1):
  for j in range(1,lastcol+1):
   col=L(j);src_ref=f"'{srcsh}'!{col}{r}";guard=f"'{srcsh}'!A{r}=\"\"" if srcsh=='Clients' else f'COUNTA(\'{srcsh}\'!A{r}:H{r})=0'
   value(master,top+3+r-6,j,f'=IF(OR({guard},ISBLANK({src_ref})),"",{src_ref})',fmt=wb[srcsh].cell(r,j).number_format)
for rr in range(6,E+1):
 for j in set(range(1,22))-{5}:master.cell(rr,j).protection=unlock
view_lock(master);master['A360'].comment=Comment('All periods, including entries without a usable period. The current month filter never hides a row here.','Basma / Operations')
# Rework breakdown complements the existing four charts, without replacing one.
ls['Y1']='First blocking category';ls['Z1']='Held records'
for r,lab in enumerate(['Identity / site','Time / calendar','Rates / adjustments','Review evidence'],2):
 ls[f'Y{r}']=lab;ls[f'Z{r}']=f'=COUNTIFS({rng(P,"V")},1,{rng(P,"AK")},Y{r})'
band(d,64,'04    /    Time-to-pay controls',2,16)
merge(d,65,2,8,'Why payroll records are held',INK,WHITE,15,True);merge(d,65,10,16,'Transfer & rework',INK,WHITE,15,True);d.row_dimensions[65].height=27
merge(d,66,2,8,'ONE PRIMARY BLOCKER PER HELD LINE   /   CURRENT SCOPE',MUT,WHITE,9,True);merge(d,66,10,16,'CORRECT AT SOURCE   /   NO DUPLICATE PAYROLL ENTRY',MUT,WHITE,9,True);d.row_dimensions[66].height=19
flowchart=profile_chart(1,8,25,26,5,['BA4952','D5A427','8FA6A5','66998E']);flowchart.anchor=TwoCellAnchor(_from=AnchorMarker(col=1,row=66),to=AnchorMarker(col=8,row=79))
for rr in range(67,81):d.row_dimensions[rr].height=14
for rr,label,f in [(67,'IMPORTED TO PAYROLL',f'=B20&" of "&SUM({rng(T,"BE")})&" scoped time lines"'),(71,'HELD FOR REVIEW','=F20&" lines  /  SAR "&TEXT(F15,"#,##0.00")'),(75,'CORRECTED & RECHECKED',f'=COUNTIFS({rng(P,"V")},1,{rng("Audit","S")},"Remade")&" remade lines"')]:
 merge(d,rr,10,16,label,MUT,PALE,9,True);merge(d,rr+1,10,16,f,INK,WHITE,16,True);d.row_dimensions[rr+1].height=25
merge(d,79,10,16,'Open all-period source register  →',TEAL,MINT,10,True);jump(d,'J79','Main Dataset','A360','Open all-period source register  →')
merge(d,81,2,16,f'=IF(F20=0,"No held payroll lines in this scope.","FIRST BLOCKER   /   "&INDEX(\'Slicer Data\'!Y2:Y5,MATCH(MAX(\'Slicer Data\'!Z2:Z5),\'Slicer Data\'!Z2:Z5,0))&" accounts for "&MAX(\'Slicer Data\'!Z2:Z5)&" of "&F20&" held lines.")',MUT,PALE,11,True);d.row_dimensions[81].height=25
merge(d,139,2,8,'Time-to-pay controls  →',TEAL,MINT,10,True);jump(d,'B139','Dashboard','B64','Time-to-pay controls  →');merge(d,139,10,16,'Source hub & audit  →',TEAL,MINT,10,True);jump(d,'J139','Main Dataset','A360','Source hub & audit  →');d.row_dimensions[139].height=22
merge(d,7,2,10,'="REPORT WINDOW   /   "&TEXT(\'Control Center\'!B65,"dd mmm yyyy")&" – "&TEXT(\'Control Center\'!B66,"dd mmm yyyy")&"   •   "&\'Control Center\'!B67&" calendar days"',MUT,PALE,9,True)
merge(d,7,11,16,'Date controls & next period  →',TEAL,MINT,9,True);jump(d,'K7','Control Center','A64','Date controls & next period  →');d.row_dimensions[7].height=19
# Link the upper workforce panel to the relocated ranking without copying rows.
cell=merge(d,34,2,16,'View top ten employees for the selected ranking  →',TEAL,MINT,10,True);jump(d,'B34','Dashboard','B121','View top ten employees for the selected ranking  →');d.row_dimensions[34].height=20
d.print_area='A1:Q142'

# Put the stable review record and its source within one click of the moving queue.
for r in range(6,N+1):
 audit[f'J{r}']=f'=IF($P{r}="","",HYPERLINK("#\'Audit\'!V"&($P{r}+5),"Review →"))'
audit['J5']='Rework log';audit.column_dimensions['J'].width=14
# Print controls: split the overview at section boundaries, never through a KPI card.
d.page_setup.orientation='portrait';d.page_setup.paperSize=d.PAPERSIZE_A3;d.page_setup.fitToHeight=0;d.row_breaks.append(Break(id=44));d.row_breaks.append(Break(id=81));d.row_breaks.append(Break(id=120));d.print_options.horizontalCentered=True
# Fixed capacity blank lines are input space, not placeholder data. Only populated data prints by default.
for sh,area in [('Main Dataset','A1:U208'),(M,'A1:X208'),('Clients','A1:O27'),(R,'A1:O208'),('Expense Manage','A1:J125'),('Audit','A1:O40')]:wb[sh].print_area=area
for s in wb:
 s.sheet_view.tabSelected=(s==d)
 if s!=d and s.sheet_state=='visible':s.sheet_view.zoomScale=85
# Consolidate four redundant navigation surfaces without deleting their records.
# Employee attributes already occupy identical rows on Main Dataset. Retain
# their checks and compact the per-person coverage report into that same row.
from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string as colnum
from openpyxl.worksheet.table import Table, TableStyleInfo
employee_view=wb[M];coverage_view=wb[R];statement_view=wb['Timesheet and Salarysheet'];examples=wb['timesheet Samples ']
cover_cols={1:2,2:26,3:7,4:8,5:10,6:17,7:32,8:33,9:34,10:35,11:36,12:37,13:38,14:22,15:5,16:25}
# Only local A1 operands are remapped. External source ranges, quoted text,
# named rules and constants are left untouched, including absolute references.
a1=re.compile(r'(\$?)([A-Z]{1,3})(\$?)([0-9]+)\Z')
def relocated(ref,cols,dy=0):
 parts=ref.split(':');matches=[a1.fullmatch(v) for v in parts]
 if not all(matches):return ref
 return ':'.join(m[1]+L(cols.get(colnum(m[2]),colnum(m[2])))+m[3]+str(int(m[4])+dy) for m in matches)
def local_formula(f,cols,dy=0):
 tok=Tokenizer(f)
 for tkn in tok.items:
  if tkn.type=='OPERAND' and tkn.subtype=='RANGE' and '!' not in tkn.value:tkn.value=relocated(tkn.value,cols,dy)
 return '='+''.join(x.value for x in tok.items)
for j in range(22,32):
 hdr(master,5,j,employee_view.cell(5,j).value);master.column_dimensions[L(j)].width=22
 for r in range(6,E+1):
  src_cell=employee_view.cell(r,j);dest=master.cell(r,j,src_cell.value);dest._style=copy.copy(src_cell._style)
for j in range(7,14):
 col=cover_cols[j];hdr(master,5,col,coverage_view.cell(5,j).value);master.column_dimensions[L(col)].width=19
 for r in range(6,E+1):
  src_cell=coverage_view.cell(r,j);dest=master.cell(r,col,local_formula(src_cell.value,cover_cols));dest._style=copy.copy(src_cell._style)
hide(master,'Y','AE');master.auto_filter.ref=f'A5:AL{E}';jump(master,'B3','Main Dataset','AF5','COVERAGE  →')
# Preserve the directory search and return its original name/ID/assignment.
value(master,4,2,'Find exact ID / name');value(master,4,3,None,True,'@').protection=unlock
search=employee_view['E3'].value.replace('D3','C4');merge(master,4,4,9,search,TEAL,MINT,10);master.row_dimensions[4].height=25
merge(master,4,10,13,'Employment & document checks  →',TEAL,MINT,10,True);jump(master,'J4','Main Dataset','V5','Employment & document checks  →')
for label,bg,fc in [('Expired',ROSE,BAD_TEXT),('Due soon',YELLOW,WARN_TEXT),('Current',MINT,OK_TEXT)]:flag(master,f'V6:W{E}',f'V6="{label}"',bg,fc)
ink_rule(master,f'V6:W{E}','V6="Missing date"',MUT);flag(master,f'X6:X{E}','AND(X6<>"",X6<>"OK")');flag(master,f'AG6:AG{E}','OR(AG6="On hold",AG6="Duplicate")');flag(master,f'AG6:AG{E}','AG6="Ready"',MINT,OK_TEXT)
# The client statement is a filtered output of Salarysheet, not a second pay
# engine. Place its totals at A:P and daily marks directly below the ledger's
# matching day 01–31 columns. The original All / Ready / Hold selector remains.
shift=430;stmt_cols={**{j:j for j in range(1,6)},**{j:j+32 for j in range(6,37)},**{j:j-31 for j in range(37,48)},48:81}
merge(pay,431,1,16,'Client payroll statement',WHITE,NAVY,20,True);pay.row_dimensions[431].height=36
merge(pay,432,1,16,'FILTERED STATEMENT    /    Follows all dashboard filters. Select a release state; scroll right to see the 31 daily marks.',MUT,PALE,11);pay.row_dimensions[432].height=25
jump(pay,'A433','Dashboard','B8','FILTERS  →');jump(pay,'B433',T,'A5','TIME SOURCE  →');value(pay,433,3,'Release state  →');value(pay,433,4,'All',True).protection=unlock;dv(pay,'D433','list',arr(['All','Ready','Hold']),blank=False)
value(pay,433,5,'Ready net →');merge(pay,433,6,8,'=SUM(O436:O855)',GREEN,MINT,12,True).number_format=MONEY;value(pay,433,9,'Rows →');pay['J433']='=COUNT(CC436:CC855)';pay.row_dimensions[433].height=24
for j in range(1,49):
 col=stmt_cols[j];hdr(pay,435,col,statement_view.cell(5,j).value)
 for r in range(6,N+1):
  src_cell=statement_view.cell(r,j);f=local_formula(src_cell.value,stmt_cols,shift).replace('ROW()-5,','ROW()-435,');dest=pay.cell(r+shift,col,f);dest._style=copy.copy(src_cell._style)
merge(pay,434,38,68,'ATTENDANCE    /    Source day 01–31 for the selected statement rows',WHITE,TEAL,10,True);pay.row_dimensions[435].height=32
jump(pay,'F3',P,'A431','STATEMENT  →');jump(pay,'K433',P,'AL435','DAY MARKS  →');jump(pay,'AL433',P,'A431','SUMMARY  →');jump(pay,'A430',P,'A1','↑  LEDGER');hide(pay,'CC','CC');flag(pay,'E436:E855','E436="Hold"');flag(pay,'E436:E855','E436="Ready"',MINT,GREEN)
# A second independent autofilter for the report, separate from the ledger.
stmt_table=Table(displayName='ClientStatement',ref='A435:P855');stmt_table.tableStyleInfo=TableStyleInfo(name='TableStyleMedium4',showFirstColumn=False,showLastColumn=False,showRowStripes=False,showColumnStripes=False);pay.add_table(stmt_table)
nm('ClientStatementSummary',"'Salarysheet'!$A$431:$P$855");jump(wb[W],'B3',P,'A431','CLIENT STATEMENT  →')
# Keep the complete October reference examples (including weekday formulas
# and merged headings) under the guide. They never become live payroll lines.
archive_shift=122;band(g,120,'09    /    October examples · reference only');merge(g,121,1,8,'Original sample layouts and entries are preserved below; none are included in attendance, pay or dashboard totals.',MUT,PALE)
for c in examples._cells.values():
 if c.value is None:continue
 dst=g.cell(c.row+archive_shift,c.column);dst.value=Translator(c.value,origin=c.coordinate).translate_formula(dst.coordinate) if c.data_type=='f' else c.value;dst._style=copy.copy(c._style)
 if c.comment:dst.comment=copy.copy(c.comment)
for area in examples.merged_cells.ranges:g.merge_cells(start_row=area.min_row+archive_shift,end_row=area.max_row+archive_shift,start_column=area.min_col,end_column=area.max_col)
for r in range(1,17):
 if examples.row_dimensions[r].height:g.row_dimensions[archive_shift+r].height=examples.row_dimensions[r].height
# Reference migration: same employee rows and IDs, one authoritative directory.
# Coverage references refer to the newly integrated fields in that same row.
removed={M,R,statement_view.title,examples.title};cover_ref=re.compile(r"'Timesheet'!(\$?[A-P]\$?\d+(?::\$?[A-P]\$?\d+)?)")
def retarget(expr):
 expr=expr.replace("'Employees'!","'Main Dataset'!").replace("'Timesheet and Salarysheet'!$D$3","'Salarysheet'!$D$433")
 return cover_ref.sub(lambda m:"'Main Dataset'!"+relocated(m[1],cover_cols),expr)
for sh in list(wb):
 if sh.title in removed:continue
 for c in sh._cells.values():
  if c.data_type=='f' and isinstance(c.value,str):c.value=retarget(c.value)
  if c.hyperlink and c.hyperlink.location:
   for old,target in [(M,"'Main Dataset'!A4"),(R,"'Main Dataset'!AF5"),(statement_view.title,"'Salarysheet'!A431"),(examples.title,"'Control Center'!A120")]:
    if c.hyperlink.location.startswith("'"+old+"'!"):c.hyperlink.location=target
for n in wb.defined_names.values():n.attr_text=retarget(n.attr_text)
for title in removed:wb.remove(wb[title])
order=['Dashboard','Main Dataset',T,P,'Audit','Clients',W,'Expense Manage','Control Center','Slicer Data'];wb._sheets=[wb[title] for title in order];wb.active=0;M=master.title
# Short tab guide records the dependency decision for every retained tab and
# the new home of every consolidated report. Keep original data lineage notes.
g['B22']='Main Dataset is the sole employee directory and source of IDs, salaries, document dates and assignments. Its coverage section adds the selected month’s time and pay summaries; Clients holds the contracts.'
g['B28']='Salarysheet → Client statement follows the dashboard filters and a Ready / Hold selector. Main Dataset → Coverage shows whether each IQAMA has a time entry for the selected month.'
g['B53']='Nine operational tabs and one hidden calculation helper remain. The searchable directory and roster coverage now share Main Dataset. Find the filtered client statement on Salarysheet, and the original examples below this guide.'
band(g,100,'08    /    Workbook map · nine working tabs, one formula helper');jump(g,'E3',g.title,'A100','WORKBOOK MAP');jump(g,'F3',g.title,'A120','EXAMPLES  →')
for row,(sh,desc) in enumerate([('Dashboard','Filtered workforce, compliance, pay and rework KPIs; charts and ranked records.'),(M,'Single employee master, exact-name / IQAMA lookup, expiry checks and per-ID roster coverage.'),(T,'Only attendance input. Validates each submitted row against the master and client rules.'),(P,'Calculates settlement, reflects daily marks and includes the filtered client statement at row 431.'),('Audit','Payroll holds, employment exceptions, imported evidence and signed corrections.'),('Clients','Contractual rate cards and client-specific absence and overtime rules.'),(W,'Site coverage, monthly yearbook and ranked chart inputs; reconciles site counts to the current scope.'),('Expense Manage','Dated source claims and approvals. Company-wide costs feed the yearbook independently of payroll.'),('Control Center','Operating settings, date helpers, release rules, workbook map and archived examples.'),('Slicer Data','Hidden formula helper for dropdown lists, ranks and the five charts. Not an input surface.')],102):
 value(g,row,1,sh);merge(g,row,2,8,desc,INK,PALE if row%2 else WHITE);g.row_dimensions[row].height=25
 if sh!='Slicer Data':jump(g,f'A{row}',sh,'A1',sh+'  →')
guide(113,'Directory & roster coverage','Former Employees and Timesheet tabs → Main Dataset. Use the search on row 4 or Coverage on row 3. Each employee’s submission, days and pay are on the same row as their master ID.')
guide(114,'Filtered client statement','Former Timesheet and Salarysheet tab → Salarysheet, row 431. Use Statement on row 3. Filter results by release state; day marks are in AL:BP. To print this report alone, set the populated statement range as the print selection.')
guide(115,'Examples','Former timesheet Samples tab → this guide, row 120. Examples are kept for reference only; they do not feed the live time register.')
g.print_area='A1:H115';jump(d,'N136',M,'A4','Employee directory  →');jump(d,'J136',M,'AF5','Roster coverage  →');jump(d,'F136',P,'A431','Client statement  →')
# Four compact, non-overlapping-in-purpose measures; only references to the
# existing filtered roster chart data. No change to any signed payroll rule.
small_metrics=[('ACTIVE WORKFORCE SHARE','=IFERROR(B108/J20,"—")',PCT,OK_TEXT,'Active employees ÷ current workforce in scope.'),('EXPIRED IQAMAS · SHARE','=IFERROR(J108/J20,"—")',PCT,BAD_TEXT,'Expired IQAMAs ÷ current workforce in scope, including missing dates.'),('IQAMA · OVER 60 DAYS',"='Slicer Data'!T6",INT,OK_TEXT,'Valid dates more than 60 days from the compliance date.'),('IQAMA · MISSING DATE',"='Slicer Data'!T7",INT,WARN_TEXT,'In-scope employees without a usable IQAMA expiry date.')]
for c,(label,f,fmt,co,tip) in zip(starts,small_metrics):
 merge(d,28,c,c+2,label,MUT,PALE,9,True);v=merge(d,29,c,c+2,f,INK,PALE,17,True);v.number_format=fmt;v.comment=Comment(tip,'Basma / Operations');ink_rule(d,v.coordinate,f'AND(ISNUMBER({v.coordinate}),{v.coordinate}>0)',co)
d.row_dimensions[28].height=18;d.row_dimensions[29].height=28
# Client-first master: swap the two original fields, not their values in a
# duplicate column. All live references (including external ranges) follow
# the client from G to A. The supplied source sequence remains available in G.
for rr in range(5,E+1):
 ca,cg=master.cell(rr,1),master.cell(rr,7);av,gv=ca.value,cg.value;sa,sg=copy.copy(ca._style),copy.copy(cg._style);ca.value=gv;cg.value=av;ca._style=sg;cg._style=sa
master['A5']='Client / worksite';master.column_dimensions['A'].width=23;master.column_dimensions['G'].width=20
# Group complete master rows once by client and then employee, keeping all
# row-local calculated fields in place. Only employee input cells are reordered.
src_inputs=[j for j in range(1,22) if j!=5];staff_rows=[[master.cell(rr,j).value for j in src_inputs] for rr in range(6,E+1) if master[f'B{rr}'].value]
for rr,vs in enumerate(sorted(staff_rows,key=lambda v:(str(v[0]).casefold(),str(v[1]).casefold())),6):
 for j,v in zip(src_inputs,vs):master.cell(rr,j).value=v
for rr in range(6,E+1):
 for j in range(22,39):
  c=master.cell(rr,j)
  if c.data_type=='f':c.value=local_formula(c.value,{1:7,7:1})
master['D4']=local_formula(master['D4'].value,{1:7,7:1})
# Limit reference rewrites to employee rows, never to the distinct lower hub tables.
clientref=re.compile(r"'Main Dataset'!(?P<rng>\$?[AG]\$?\d+(?::\$?[AG]\$?\d+)?)\b")
def clientfirst(expr):
 def sub(m):
  if not all(6<=int(n)<=E for n in re.findall(r'\d+',m['rng'])):return m[0]
  return "'Main Dataset'!"+relocated(m['rng'],{1:7,7:1})
 return clientref.sub(sub,expr)
for sh in wb:
 for c in sh._cells.values():
  if c.data_type=='f' and "'Main Dataset'!" in c.value:c.value=clientfirst(c.value)
for n in wb.defined_names.values():n.attr_text=clientfirst(n.attr_text)
# Compact one-row-per-client map. No duplicate client tabs and no second
# employee master: each row drills into the existing registers and reports.
index_top=1100;band(master,index_top,'Client workspace · one row per rate card or roster-only site',1,12)
merge(master,1101,1,12,'Choose a client in column A. Links open the matching client record where available; Dashboard, rules and the archive are shared views.',MUT,PALE,10)
for j,lab in enumerate(['Client / worksite','Main dataset','Dashboard','Time input','Payroll','Audit','Clients','Site analysis','Expenses','Control center','Master staff','Time lines'],1):hdr(master,1102,j,lab)
for rr in range(6,C+1):
 outrow=1103+rr-6;site=f'$A{outrow}';value(master,outrow,1,f'=IF(\'Clients\'!A{rr}="","",\'Clients\'!A{rr})')
 def btn(j,tab,addr,label):value(master,outrow,j,f'=IF({site}="","",HYPERLINK("#\'{tab}\'!"&{addr},"{label}"))')
 btn(2,M,f'IFERROR("B"&(MATCH({site},{rng(M,"A",E)},0)+5),"A5")','Roster');btn(3,'Dashboard','"F9"','Site filter');btn(4,T,f'IFERROR("B"&(MATCH({site},{rng(T,"B")},0)+5),"B5")','Intake');btn(5,P,f'IFERROR("C"&(MATCH({site},{rng(P,"B")},0)+5),"A5")','Ledger');btn(6,'Audit',f'IFERROR("F"&(MATCH({site},{rng("Audit","E")},0)+5),"D3")','Review');btn(7,'Clients',f'"A{rr}"','Rate card');btn(8,W,f'"A{rr}"','Scorecard');btn(9,'Expense Manage',f'IFERROR("C"&(MATCH({site},{rng("Expense Manage","C",X)},0)+5),"C5")','Claims');btn(10,'Control Center','"A64"','Date rules')
 value(master,outrow,11,f'=IF({site}="","",COUNTIF({rng(M,"A",E)},{site}))',fmt=INT);value(master,outrow,12,f'=IF({site}="","",COUNTIF({rng(T,"B")},{site}))',fmt=INT)
jump(master,'H3',M,'A1100','Client workspace →');master['A2']='Client first · one employee per row. Filter column A for a roster. Client workspace opens one row per client with links to the working tabs.'
# Keep historical originals, but explicitly detach them from new paste slots.
# The live Register/Payroll/Audit must start empty, not calculate pay for
# imported records. Keep the 73 received rows out of all live formula ranges.
band(audit,430,'Received September evidence · archive only / excluded from live payroll',1,14)
jump(audit,'O430','Audit','A1','↑ Live review');jump(audit,'B428','Audit','A430','Received archive →')
merge(audit,431,1,14,'Original imported submissions and their findings. The live time register opens empty. These records are reference evidence, not new paste slots.',MUT,PALE,10)
archive_headers=['Imported finding']+[str(tim.cell(5,j).value) for j in rawcols]
for j,title in enumerate(archive_headers,1):hdr(audit,432,j,title)
for rr in range(6,79):
 value(audit,433+rr-6,1,audit[f'AA{rr}'].value)
 for k,j in enumerate(rawcols):
  cell=audit.cell(rr,snap0+k);value(audit,433+rr-6,k+2,cell.value,fmt=cell.number_format);cell.value=None
 audit[f'AA{rr}'].value=None
for rr in range(6,N+1):
 for j in [*rawcols,67]:tim.cell(rr,j).value=None
 # The time row exists only when data has actually been pasted; no imported
 # row number or previous document mark can activate a blank payroll slot.
 tim[f'BN{rr}']=f'=--(COUNTA(A{rr}:D{rr},H{rr}:AL{rr},AR{rr}:AT{rr})>0)'
 # Disallow typed numeric text in regular-hour input; reject an unfinished
 # submission through the existing verification checks rather than guessing.
 # Source-origin metadata belongs to the archived import, not new rows.
 tim[f'AV{rr}'].protection=xl.styles.Protection(locked=True);tim[f'AW{rr}'].protection=xl.styles.Protection(locked=True)
# Two deliberately separate paste blocks prevent overwriting verification
# formulas. Existing month/client/day/adjustment validation stays on the cells.
tim['A1']='Timesheet input';tim['A2']='Paste values only: A:D = month start, client, full name, IQAMA; H:AL = days 1–31. Leave E:G for verification. Each employee needs one row per month.'
tim['A4']='Paste →';style(tim['A4'],TEAL,MINT,10);merge(tim,4,2,4,'1   Identity  /  A:D',BLUE,SKY,10);merge(tim,4,5,7,'2   Automatic master check',MUT,PALE,10);merge(tim,4,8,16,'3   Daily entries  /  H:AL  ·  P, A, OFF, 0.5P or numeric units',BLUE,SKY,10)
dv(tim,f'D6:D{N}','custom','IFERROR(AND(LEN(TRIM(D6&""))=10,TEXT(VALUE(TRIM(D6&"")),"0000000000")=TRIM(D6&"")),FALSE)',msg='Use the exact 10-digit IQAMA from Main Dataset. Paste values only.')
# Native Excel filter arrows really filter the input register. The quick view
# offers automatic client/release filters without a macro or changing row IDs.
tim.auto_filter.ref=f'A5:BS{N}';tim.protection.autoFilter=False
jump(tim,'E3',T,'A430','Filter entries ▾');tim['E3'].comment=Comment('Open the live filter panel. To filter the editable rows in place, use the arrows in row 5.','Basma / Operations')
filter_start,filter_end=435,854
band(tim,430,'Filter entries · read-only results',1,7)
merge(tim,431,1,7,'Pick a client and release status. All periods are included. Open time row returns to the original paste row. Reset both selectors to All to show every entry.',MUT,PALE,10)
value(tim,432,1,'Client');value(tim,432,2,'All',True).protection=unlock;value(tim,432,3,'Release');value(tim,432,4,'All',True).protection=unlock
value(tim,432,5,'Matches');value(tim,432,6,f'=COUNT(BT6:BT{N})',fmt='0');jump(tim,'G432',T,'A6','Back to paste →');dv(tim,'B432','list','Sites',blank=False);dv(tim,'D432','list',arr(['All','Ready','Hold']),blank=False)
for j,lab in enumerate(['Time row','Client / site','Employee','IQAMA','Identity check','Release','First finding','Period','Days logged'],1):hdr(tim,434,j,lab)
for rr in range(6,N+1):
 tim[f'BT{rr}']=f'=IF(AND(BN{rr}=1,OR($B$432="All",B{rr}=$B$432),OR($D$432="All",BQ{rr}=$D$432)),COUNT($BT$5:BT{rr-1})+1,"")'
 outrow=filter_start+rr-6;ix=f'$BU{outrow}';tim[f'BU{outrow}']=f'=IFERROR(MATCH(ROW()-434,$BT$6:$BT${N},0),"")'
 value(tim,outrow,1,f'=IF({ix}="","",HYPERLINK("#\'Timesheet_Input\'!A"&({ix}+5),"Time row "&({ix}+5)))')
 for j,src_col in enumerate(['B','C','D','E','BQ','BR','A','AP'],2):
  ref=f'INDEX(${src_col}$6:${src_col}${N},{ix})';value(tim,outrow,j,f'=IF({ix}="","",IF(ISBLANK({ref}),"",{ref}))',fmt=PER if j==8 else INT if j==9 else None)
hide(tim,'BT','BU');flag(tim,'F435:F854','F435="Hold"');flag(tim,'F435:F854','F435="Ready"',MINT,GREEN)
# First live claims/roster inputs stay intact; archive is a separate read-only
# time record. Correct any earlier language implying that the feed is prefilled.
audit['A2']='Live holds · correct pasted entries at their time source. The rework log stays attached to fixed time rows; original September submissions are kept separately in the archive.'
jump(audit,'G3','Audit','A430','Received archive →');audit['AE4']='Live evidence fields · the original September submissions are filed in the received archive below, not in these live slots.'
g['B23']='Start with the blank Timesheet_Input register. Paste values into A:D, then H:AL, starting on the same row. Checks run against Main Dataset and Clients immediately. Never paste over calculated columns E:G.'
g['B48']='The 73 original September submissions are retained in Audit → Received archive. The earlier 50-row input overlaps those records. Neither historical set populates a new live paste slot.'
g['B58']='Historical build check only: 73 supplied records, 55 passing and 18 held; ready SAR 75,630.06 + held SAR 23,163.68 = estimated SAR 98,793.74. The live register starts empty and only pasted entries create live payroll.'
g['B75']='The received archive preserves the original 73 submitted rows and their initial findings independently of the live input. Record review notes before correcting a new suspect line. Keep a dated copy for closed periods.'
# Compact, regular-weight section titles; table header rows remain distinct.
# Status formatting, financial controls, charts and date logic are untouched.
for sh in (master,tim):
 sh.sheet_format.defaultRowHeight=15;sh.sheet_view.zoomScale=85
 style(sh['A1'],WHITE,NAVY,17,False);sh.row_dimensions[1].height=27;sh.row_dimensions[2].height=20;sh.row_dimensions[3].height=19;sh.row_dimensions[4].height=17;sh.row_dimensions[5].height=22
for sh,rows in [(master,[360,790,840,1100]),(tim,[430]),(audit,[430])]:
 for rr in rows:style(sh.cell(rr,1),WHITE,TEAL,11,False);sh.row_dimensions[rr].height=19
for j in range(8,39):tim.column_dimensions[L(j)].width=8
for sh in wb:
 if sh not in (master,tim):continue
 for c in sh[3]:
  if c.hyperlink:style(c,TEAL,MINT,9,False)
master['A1100'].comment=Comment('Client-first navigation and employee roster use the same existing client and employee sources. No duplicate client sheets are created.','Basma / Operations')

# One compact case register, fed by explicit case records only. Keep the staff
# master and attendance history intact; an absence is never a runaway finding.
case_tab='Runaway Records';cases=wb.create_sheet(case_tab,wb.index(audit)+1);order=wb.sheetnames
setup(cases);cases.sheet_view.zoomScale=90;cases.sheet_format.defaultRowHeight=15;cases.sheet_properties.tabColor=TEAL
for col,wi in [('A',25),('B',20),('C',38),('D',18),('E',18)]:cases.column_dimensions[col].width=wi
merge(cases,1,1,5,'Runaway records',WHITE,NAVY,17,False);cases.row_dimensions[1].height=27
merge(cases,2,1,5,'Reported cases only · one record per employee. Select an employee name to update the source. Resolved records stay listed.',MUT,PALE,10);cases.row_dimensions[2].height=20
jump(cases,'A3',M,'AM5','Case fields →');jump(cases,'B3','Dashboard','A1','Dashboard →');jump(cases,'C3','Control Center','A116','Register guide →');value(cases,3,4,'Records');value(cases,3,5,f'=COUNT(F6:F{E})',fmt='0');cases.row_dimensions[3].height=18
merge(cases,4,1,5,'Month is the reported month, not the joining or payroll date. “Not recorded” means the source did not supply it.',MUT,PALE,9);cases.row_dimensions[4].height=18
for j,h in enumerate(['Client','Month / year','Employee','IQAMA','Case status','Source index'],1):style(cases.cell(5,j,h),WHITE,TEAL,10,False)
cases.row_dimensions[5].height=20;cases.auto_filter.ref=f'A5:E{E}';cases.freeze_panes='B6';cases.print_title_rows='1:5';cases.print_area=f'A1:E{E}';view_lock(cases)
for j,(label,width) in enumerate([('Runaway case status',21),('Reported month / year',23),('Client at report',25),('Case sequence',14)],39):
 style(master.cell(5,j,label),WHITE,TEAL if j<=41 else NAVY,10,False);master.column_dimensions[L(j)].width=width
merge(master,4,39,41,'Case fields · use the documented client and report month; retain Resolved records.',MUT,PALE,9)
case_states=['Reported','Under review','Resolved'];origin_flags={'RUNAWAY','RAN AWAY','RUN AWAY'}
for rr in range(6,E+1):
 is_marked=any(str(master.cell(rr,j).value or '').strip().upper() in origin_flags for j in (1,10,20,21))
 for j,v in [(39,'Reported' if is_marked else None),(40,None),(41,None)]:value(master,rr,j,v,True,PER if j==40 else None).protection=unlock
 # A whole-field report marker is intentional. Unrelated prose, suspended,
 # terminated, a missing shift or a missing pay row must never open a case.
 tests=[f'AM{rr}="{state}"' for state in case_states]+[f'TRIM({k}{rr})="{mark}"' for k in ('A','J','T','U') for mark in sorted(origin_flags)]
 master[f'AP{rr}']=f'=IF(AND(B{rr}<>"",OR({",".join(tests)})),COUNT($AP$5:AP{rr-1})+1,"")'
 ix=f'$F{rr}';cases[f'F{rr}']=f'=IFERROR(MATCH(ROW()-5,{rng(M,"AP",E)},0),"")'
 def pick(col):return f'INDEX({rng(M,col,E)},{ix})'
 reported=pick('AO');assignment=pick('A');mo=pick('AN');state=pick('AM');real_site=lambda ref:f'AND({ref}<>"",{ref}<>"-",{ref}<>"Runaway",{ref}<>"Ran away",{ref}<>"Run away",{ref}<>"ZZZ INACTIVE",{ref}<>"VACATION")'
 fs={1:f'IF({real_site(reported)},{reported},IF({real_site(assignment)},{assignment},"Not recorded"))',2:f'IF({mo}="","Not recorded",IF(IFERROR(AND(ISNUMBER({mo}),{mo}>0,{mo}=INT({mo}),DAY({mo})=1),FALSE),{mo},"Check month"))',3:f'HYPERLINK("#\'Main Dataset\'!AM"&({ix}+5),{pick("B")})',4:f'IF({pick("C")}="","Not recorded",{pick("C")}&"")',5:f'IF(OR({state}="Under review",{state}="Resolved"),{state},"Reported")'}
 for j,body in fs.items():value(cases,rr,j,f'=IF({ix}="","",{body})',fmt=PER if j==2 else None)
for state,co in zip(case_states,[BAD_TEXT,WARN_TEXT,OK_TEXT]):
 for area in (f'A6:A{E}',f'E6:E{E}'):ink_rule(cases,area,f'$E6="{state}"',co)
ink_rule(cases,f'B6:B{E}','B6="Check month"',BAD_TEXT);ink_rule(cases,f'B6:B{E}','B6="Not recorded"',MUT)
for c in cases[3]:
 if c.hyperlink:style(c,TEAL,MINT,9,False)
hide(cases,'F','F');hide(master,'AP','AP');master.auto_filter.ref=f'A5:AO{E}'
dv(master,f'AM6:AM{E}','list',arr(case_states),msg='Use a documented case status. Choose Resolved to retain a closed case.')
dv(master,f'AN6:AN{E}','custom','IFERROR(AND(ISNUMBER(AN6),AN6=INT(AN6),AN6>0,DAY(AN6)=1),FALSE)',msg='Use the first day of the documented month. Leave blank if the month is unknown.')
jump(master,'I3',case_tab,'A5','Runaway records →');jump(audit,'H3',case_tab,'A5','Runaway records →');jump(d,'B140',case_tab,'A5','Runaway records →');d.row_dimensions[140].height=18
jump(g,'A112',case_tab,'A1','Runaway Records →');merge(g,112,2,8,'Client-first register of explicitly reported or resolved cases. Select an employee to edit the case fields on Main Dataset.',MUT,PALE,10);g.row_dimensions[112].height=23
g['A100']='08    /    Workbook map · ten working tabs, one formula helper';g['B53']='Ten operational tabs and one hidden calculation helper. Main Dataset holds the employee directory and case inputs; Runaway Records shows explicit cases. Client statements remain on Salarysheet; examples are archived below.'
guide(116,'Runaway register','Use Main Dataset AM:AO for case status, month and client at report. An exact runaway marker in the source is also listed. Red = reported, amber = under review, green = resolved. Records are linked, not removed from the master; no payroll rule is changed.')
guide(117,'Missing case details','Enter only a documented month. If Client at report is blank, the register shows a usable master site, otherwise Not recorded. Enter the historical client before reassignment to retain it. Do not delete the case or master row when resolving a report.')
g.print_area='A1:H117';master['AM5'].comment=Comment('Record a documented case only. A status of Suspended or Terminated by itself does not create a case. Keep Resolved to retain the record.','Basma / Operations')
# Keep the same four-field paste block and put client, month/year first on
# each intake row. Retarget live A/B references only within the intake grid;
# leave the archived evidence, chart scopes and lower quick-filter controls alone.
for rr in range(5,N+1):
 a,b=tim.cell(rr,1),tim.cell(rr,2);a.value,b.value=b.value,a.value;a._style,b._style=copy.copy(b._style),copy.copy(a._style)
tim['A5']='Client / worksite';tim['B5']='Month / year';tim.column_dimensions['A'].width=22;tim.column_dimensions['B'].width=17
# All time / month and client ranges in other sheets reference the same source.
remote_ref=re.compile(r"'Timesheet_Input'!(?P<rg>\$?[AB]\$?\d+(?::\$?[AB]\$?\d+)?)\b")
def input_ref(rg):
 parts=rg.split(':');matches=[a1.fullmatch(p) for p in parts]
 if not all(matches) or not all(6<=int(m[4])<=N for m in matches):return rg
 if len(parts)==2 and matches[0][2]!=matches[1][2]:return rg # rectangle A:D already contains both
 return relocated(rg,{1:2,2:1})
for sh in wb:
 for c in sh._cells.values():
  if c.data_type!='f' or not isinstance(c.value,str):continue
  if sh is tim and re.search(r'(?<![\w!])\$?[AB]\$?\d+',c.value):
   toks=Tokenizer(c.value)
   for tok in toks.items:
    if tok.type=='OPERAND' and tok.subtype=='RANGE' and '!' not in tok.value:tok.value=input_ref(tok.value)
   c.value='='+''.join(tok.value for tok in toks.items)
  if "'Timesheet_Input'!" in c.value:c.value=remote_ref.sub(lambda m:"'Timesheet_Input'!"+input_ref(m['rg']),c.value)
for v in tim.data_validations.dataValidation:
 if str(v.sqref)==f'A6:A{N}':v.sqref=f'B6:B{N}';v.formula1=v.formula1.replace('A6','B6')
 elif str(v.sqref)==f'B6:B{N}':v.sqref=f'A6:A{N}'
for rr in range(1103,1143):master[f'D{rr}']=master[f'D{rr}'].value.replace('"B"&','"A"&').replace('"B5"','"A5"')
tim['A2']='Paste values only: A:D = client, month start, full name, IQAMA; H:AL = days 1–31. Leave E:G for automatic master checks. One employee per month per row.'
tim['B4']='1   Client · month · name · IQAMA';g['B23']='Start with blank Timesheet_Input. Paste client, month-start date, name and IQAMA in A:D; paste daily marks in H:AL on the same rows. E:G verifies against the master. Use Filter entries or the header arrows.'
for sh in (cases,tim):
 for c in sh[5]:
  if c.value is not None:style(c,WHITE,TEAL if sh is cases or c.column in (1,2,3,4,*range(8,39),44,45,46) else NAVY,9,True,wrap=sh is tim)

# A single, linked teaching row per working register. All examples sit outside
# the editable, counted and filtered ranges; the time register starts empty.
# Use the archived September example for the time-to-pay walkthrough. A case
# example is explicitly fictional, never an allegation against a real worker.
sample_specs={
 M:(356,'A4',41,'Master record · same employee as the time example'),
 T:(426,'F3',50,'Paste values from A:D and H:AL separately; E:G are checks'),
 P:(426,'C3',21,'Illustrative settlement for the time example · not payable'),
 'Audit':(507,'I3',9,'Archived site mismatch · compare submitted and expected sites'),
 case_tab:(357,'E4',5,'Fictional case · not a reported employee'),
 'Clients':(47,'C3',15,'Rate card used by the timesheet example'),
 W:(64,'C3',14,'One-worker illustration · not an actual site total'),
 'Expense Manage':(257,'C3',10,'Illustrative expense input · SAR 80, not a real claim'),
}
# The case sheet stays exactly five visible columns wide.
note=cases['A4'].value;cases.unmerge_cells('A4:E4');merge(cases,4,1,4,note,MUT,PALE,9)
for name,(top,shortcut,last,caption) in sample_specs.items():
 sh=wb[name];r=top+2;merge(sh,top,1,min(last-1,6),'Sample row  /  '+caption,MUT,PALE,10,False)
 back_col=min(last,7);jump(sh,L(back_col)+str(top),name,'A5','↑  Live register')
 jump(sh,shortcut,name,'A'+str(top),'Sample row →');sh.row_dimensions[top].height=18;sh.row_dimensions[top+1].height=19;sh.row_dimensions[r].height=18
 for j in range(1,last+1):
  if sh.cell(5,j).value is None:continue
  style(sh.cell(top+1,j,sh.cell(5,j).value),TEAL,'E8F2EF',9,True,'left',1,True)
 def src(col,ix):return f'INDEX({rng(M,col,E)},MATCH({ix},{rng(M,"C",E)},0))'
 def sv(j,val):
  c=sh.cell(r,j);c.value=val;c.number_format=sh.cell(6,j).number_format;style(c,INK,'F4F8F6',10,False,'left',1);return c
 if name==M:
  for j in range(1,42):
   if not 25<=j<=31:
    ref=src(L(j),"'Audit'!$E$433");sv(j,f'=IF(ISBLANK({ref}),"",{ref})')
 elif name==T:
  fields={1:3,2:2,3:4,4:5,**{7+i:5+i for i in range(1,32)}}
  for j,k in fields.items():sv(j,linked('Audit',L(k)+'433'))
  for j,formula in [(5,f'IF(AND(C{r}=F{r},A{r}=G{r}),"Matched","Review")'),(6,src('B',f'D{r}')),(7,src('A',f'D{r}')),(39,f'COUNTIF(H{r}:AL{r},"P")'),(40,f'COUNTIF(H{r}:AL{r},"OFF")'),(41,f'COUNTIF(H{r}:AL{r},"A")'),(42,f'SUM(AM{r}:AO{r})'),(43,f'IF(AP{r}=DAY(DATE(YEAR(B{r}),MONTH(B{r})+1,0)),"Complete","Review")')]:sv(j,'='+formula)
  sv(47,'Archived example only');sh[f'D{r}'].comment=Comment('Copy identity and day blocks separately using Paste Values. This example is outside all live checks and payment totals.','Basma / Operations')
 elif name==P:
  for j,k in [(1,'B'),(2,'A'),(3,'C'),(4,'D'),(5,'E'),(17,'AP'),(18,'AM'),(19,'AO'),(20,'AN')]:sv(j,linked(T,k+'428'))
  # The worked example has a monthly base, no absence, no OT and no adjustments.
  for j,f in {6:src('P',f'D{r}'),7:f'S{r}*IFERROR(INDEX({rng("Clients","E",C)},MATCH(B{r},{rng("Clients","A",C)},0)),0)',8:'0',9:'0',10:f'ROUND(F{r}+I{r},2)',11:f'MIN(F{r},G{r}+H{r})',12:f'ROUND(J{r}-K{r},2)',13:'0',14:f'L{r}',15:'"Ready"',16:'"Checks passed"',21:'0'}.items():sv(j,'='+f)
 elif name=='Audit':
  for j,k in [(2,'B'),(3,'D'),(4,'E'),(5,'C'),(6,'A')]:sv(j,linked('Audit',k+'449'))
  sv(1,'Archive row 449');sv(7,'='+src('B',f'D{r}'));sv(8,'='+src('A',f'D{r}'));sv(9,None)
  sh[f'I{r}'].comment=Comment('An archived finding is not live held pay. The example intentionally leaves this amount blank.','Basma / Operations')
 elif name==case_tab:
  for j,v in enumerate(['Example client',start,'Example employee','0000000000','Under review'],1):sv(j,v)
  sh[f'A{r}'].font=font(WARN_TEXT,10,False);sh[f'E{r}'].font=font(WARN_TEXT,10,False)
 elif name=='Clients':
  for j in range(1,16):
   ref=f'INDEX({rng("Clients",L(j),C)},MATCH(\'Timesheet_Input\'!$A$428,{rng("Clients","A",C)},0))';sv(j,f'=IF(ISBLANK({ref}),"",{ref})')
 elif name==W:
  for j,v in enumerate(["='Timesheet_Input'!A428",1,1,1,'=B66-D66','=IFERROR(D66/B66,0)',1,1,0,"='Salarysheet'!N428","='Salarysheet'!M428","='Timesheet_Input'!AM428","='Timesheet_Input'!AO428",'=IFERROR(L66/(L66+M66),0)'],1):sv(j,v)
 elif name=='Expense Manage':
  for j,v in enumerate([start,'Example employee','Example client','Travel',80,'Pending','Cash','Example travel reimbursement',f'=IF(E{r}>\'Control Center\'!$B$9,"HR review","Standard")',f'=DATE(YEAR(A{r}),MONTH(A{r}),1)'],1):sv(j,v)
# Index is a navigation aid, not another set of example transactions.
band(g,141,'10    /    One-row examples · reference only')
for rr,(name,(top,_,_,caption)) in enumerate(sample_specs.items(),142):
 jump(g,f'A{rr}',name,f'A{top}',name+'  →');merge(g,rr,2,8,caption,MUT,PALE,10);g.row_dimensions[rr].height=18
jump(g,'H3',g.title,'A141','Sample rows →');merge(d,142,2,16,'Open one-row examples  →',TEAL,MINT,10,False);jump(d,'B142',g.title,'A141','Open one-row examples  →');d.row_dimensions[142].height=18
# Size the actual table fields from rendered values, not the internal length of
# a formula. Keep headers aligned with their column: dates/day marks centered,
# amounts right-aligned and text left-aligned. Use quiet, low-height upper bands.
def fit_registers(cached):
 from math import ceil
 from datetime import date,datetime
 def width_text(v,fmt):
  if v is None:return ''
  if isinstance(v,(date,datetime)):return v.strftime('%b %Y' if fmt==PER else '%d %b %Y')
  if isinstance(v,(float,int)):
   if '%' in fmt:return f'{v:.1%}'
   if fmt==MONEY:return f'{v:,.2f}'
   return f'{v:,g}'
  return str(v)
 for name,(top,shortcut,lastcol,_) in sample_specs.items():
  sh=wb[name];cvsh=cached[name];end={M:E,T:N,P:N,'Audit':N,case_tab:E,'Clients':C,W:C,'Expense Manage':X}[name]
  sh.sheet_format.defaultRowHeight=16
  for row,size in [(1,17),(2,10)]:
   c=sh.cell(row,1);c.font=font(WHITE if row==1 else MUT,size,False);c.alignment=align('left',1,True)
  sh.row_dimensions[1].height=26;sh.row_dimensions[2].height=20;sh.row_dimensions[5].height=21
  hidden={j for cd in sh.column_dimensions.values() if cd.hidden for j in range(cd.min,cd.max+1)}
  for j in range(1,lastcol+1):
   if j in hidden:continue
   col=L(j);field=sh.cell(5,j).value
   if not field:continue
   fmt=sh.cell(6,j).number_format;values=[cvsh.cell(rr,j).value for rr in range(6,end+1)];values.append(cvsh.cell(top+2,j).value)
   day=bool(re.fullmatch(r'Day \d\d',str(field)))
   dates=fmt in (PER,DATE);nums=fmt in [MONEY,NUM,INT,PCT] or all(isinstance(v,(int,float)) for v in values if v is not None) and any(v is not None for v in values)
   halign='center' if day or dates else 'right' if nums and fmt!='@' else 'left'
   labels=[cvsh.cell(3,j).value,field];lwidth=max([len(str(v)) for v in labels if v is not None]+[0]);dw=max([len(width_text(v,fmt)) for v in values]+[0]);oldw=sh.column_dimensions[col].width
   # Multiple independent grids share a column. Never make a pre-existing
   # column narrower than an established useful width; cap new long-text growth.
   want=max(9 if day else 12,min(44,max(lwidth+3,dw+2)))
   dim=sh.column_dimensions[col];dim.width=9 if day else max(oldw,want);dim.bestFit=True
   for rr in [5,top+1]:
    h=sh.cell(rr,j);h.font=font(WHITE if rr==5 else TEAL,9,True);h.alignment=align(halign,0 if halign=='center' else 1,True)
   for rr in [*range(6,end+1),top+2]:
    c=sh.cell(rr,j);c.alignment=Alignment(horizontal=halign,vertical='center',indent=0 if halign=='center' else 1,shrinkToFit=True)
  # Fit wrapped subtitles to the real merged width, independent of sheet size.
  for row in (1,2,top):
   a=sh.cell(row,1);area=next((rg for rg in sh.merged_cells.ranges if rg.min_row==row and rg.min_col==1),None)
   if area:
    space=sum(sh.column_dimensions[L(j)].width for j in range(1,area.max_col+1));n=max(1,ceil(len(str(a.value or ''))/max(1,space*.94)));sh.row_dimensions[row].height=max(sh.row_dimensions[row].height or 16,14*n+3)
  for c in sh[3]:
   if c.hyperlink:c.font=font(TEAL,9,False);c.alignment=align('left',1,False)
 # Preserve the original grid before applying the final accessible reading style.
 for cell in ['B2','B3',*[f'B{rr}' for rr in [24,30,45,64,82,101,106,121,135]]]:
  c=d[cell];f=copy.copy(c.font);f.bold=False;f.sz=20 if cell=='B3' else 10;c.font=f;c.alignment=align('left',1)
 d.row_dimensions[3].height=31
 for sh,where in [(g,1),(g,141)]:
  c=sh.cell(where,1);f=copy.copy(c.font);f.bold=False;f.sz=17 if where==1 else 10;c.font=f;c.alignment=align('left',1);sh.row_dimensions[where].height=26 if where==1 else 18

# Final review: treat a documented departure separately from absence, leave,
# suspension and a runaway report. All counts use the same master status list.
work_states=['Active','On Leave','Suspended','Terminated','Left','Resigned'];departed_states=work_states[3:]
for val in master.data_validations.dataValidation:
 if str(val.sqref)==f'J6:J{E}':val.formula1=arr(work_states)
for rr,st in enumerate(['All',*work_states,'Unmatched'],2):ls[f'E{rr}']=st
nm('Statuses',"'Slicer Data'!$E$2:$E$9")
ls['V5']='Departed';ls['W5']='='+'+'.join(counte('J',st) for st in departed_states)
d['B32']="='Slicer Data'!W4+'Slicer Data'!W5";d['B33']='Suspended + departed; excludes leave'
# Promote the active count beside total workforce and inactive. Keep its share
# on the detailed panel: no lost metric and no duplicate headcount card.
d['B28']='TOTAL ACTIVE';d['B29']='='+counte('J','Active');d['B29'].number_format='0';d['B29'].comment=Comment('Active master employees in the current roster scope.','Basma / Operations')
d['B107']='ACTIVE WORKFORCE SHARE';d['B108']='=IFERROR(B29/J20,"—")';d['B108'].number_format=PCT;d['B109']='Active employees ÷ current roster';d['J62']=d['J62'].value.replace('B108','B29').replace(' are suspended or terminated.',' are inactive (suspended or departed).')
# A single small integer helper supplies the scoped departures count and
# the drill-down. Do not infer actual departures from notes or attendance.
master['AQ5']='Departure order';hide(master,'AQ','AQ')
for rr in range(6,E+1):
 is_departed='OR('+','.join(f'$J{rr}="{st}"' for st in departed_states)+')'
 master[f'AQ{rr}']=f'=IF(AND($Y{rr}=1,{is_departed}),COUNT($AQ$5:AQ{rr-1})+1,"")'
 exact=[f'TRIM({k}{rr})="{label}"' for k in ('A','J','T','U') for label in ['Runaway','Ran away','Run away','Left','Resigned','Terminated','Not interested']]
 master[f'AA{rr}']=f'=IF(B{rr}="","",IF(OR({",".join(exact)},AM{rr}="Reported",AM{rr}="Under review"),"Review",""))'
 # Reject a pasted zero/negative or missing control date rather than counting
 # it as a genuine expiry in 1899. Compare whole dates, not time of day.
 master[f'E{rr}']=f'=IFERROR(IF(AND(B{rr}<>"",ISNUMBER(D{rr}),D{rr}>0,ISNUMBER(ReviewDate),ReviewDate>0),INT(D{rr})-INT(ReviewDate),""),"")'
 master[f'W{rr}']=f'=IF(B{rr}="","",IF(OR(NOT(ISNUMBER(R{rr})),N(R{rr})<=0,NOT(ISNUMBER(ReviewDate)),N(ReviewDate)<=0),"Missing date",IF(INT(R{rr})<INT(ReviewDate),"Expired",IF(INT(R{rr})-INT(ReviewDate)<=InsuranceWarnDays,"Due soon","Current"))))'
for rng0 in [f'A6:A{E}',f'J6:J{E}']:
 for st in departed_states:ink_rule(master,rng0,f'AND($B6<>"",$J6="{st}")',BAD_TEXT)
ink_rule(master,f'A6:A{E}','AND($B6<>"",OR($J6="On Leave",$J6="Suspended",AND($J6="Active",$AA6="Review")))',WARN_TEXT)
ink_rule(master,f'A6:A{E}','AND($B6<>"",$J6="Active",$AA6<>"Review")',OK_TEXT)
# Eliminate empty choices in the client selectors, also when a new contract
# is entered below a gap. No filter list is separately maintained.
client_sheet=wb['Clients'];client_sheet['P5']='List order'
for rr in range(6,C+1):client_sheet[f'P{rr}']=f'=IF(A{rr}="","",COUNT($P$5:P{rr-1})+1)'
hide(client_sheet,'P','P')
for rr in range(3,43):ls[f'A{rr}']=f'=IFERROR(INDEX({rng("Clients","A",C)},MATCH(ROW()-2,{rng("Clients","P",C)},0)),"")'
num_clients=f'COUNT({rng("Clients","P",C)})'
nm('Sites',f"'Slicer Data'!$A$2:INDEX('Slicer Data'!$A$2:$A$42,1+{num_clients})")
nm('ClientNames',f"'Slicer Data'!$A$3:INDEX('Slicer Data'!$A$3:$A$42,MAX(1,{num_clients}))")
# Four additional measures use two existing spacer rows. No panel or chart
# moves; the top-ten table, sample shortcuts and all expiry cards are retained.
review_kpis=[('DEPARTED STAFF',f'COUNT({rng(M,"AQ",E)})',MUT,'Terminated, Left or Resigned. Click the heading for the scoped register.'),('ACTIVE WITHOUT TIMESHEET',f'COUNTIFS({rng(M,"Y",E)},1,{rng(M,"J",E)},"Active",{rng(M,"AF",E)},0)',WARN_TEXT,'Active master employees without an IQAMA-matched submission for the selected month.'),('BOTH DOCUMENTS CURRENT',f'COUNTIFS({rng(M,"Y",E)},1,{rng(M,"V",E)},"Current",{rng(M,"W",E)},"Current")',OK_TEXT,'Iqama and insurance are both outside their configured warning windows.'),('IQAMA  /  31–60 DAYS',"'Slicer Data'!T5",WARN_TEXT,'Non-overlapping forward band: day 31 through day 60 inclusive; not missing, expired or due today.')]
for col,(label,formula,color,definition) in zip(starts,review_kpis):
 merge(d,110,col,col+2,label,MUT,PALE,9,True);v=merge(d,111,col,col+2,'='+formula,INK,PALE,17,True);v.number_format='#,##0';v.comment=Comment(definition,'Basma / Operations');ink_rule(d,v.coordinate,f'AND(ISNUMBER({v.coordinate}),{v.coordinate}>0)',color)
d.row_dimensions[110].height=17;d.row_dimensions[111].height=25
# Existing audit archive is complete at A:AQ, rows 433–505. The superseded
# empty snapshot grid above it (AE:BT) no longer supplies pay, holds or review
# history. Drop the dead columns, not any source record, signature or archive.
audit.unmerge_cells('AE4:BT4')
for rr in range(6,N+1):
 audit[f'AC{rr}']=f'=IF(\'{T}\'!BN{rr}=0,0,--(COUNTA(V{rr}:X{rr},Z{rr})>0))'
for k in list(audit._cells):
 rr,j=k
 if (j>=31 and rr<=N) or (j in [27,28] and 5<=rr<=N):del audit._cells[k]
for col in list(audit.column_dimensions):
 if colnum(col)>43:del audit.column_dimensions[col]
assert audit.max_column==43
# Keep payroll input evidence accessible in its archive and stop showing
# non-editable legacy import fields alongside current entry fields.
hide(tim,'AV','AX');hide(pay,'BU','BW')
# Scoped departure records in Audit, not a duplicate editable roster.
# Both views read the master; missing client names are not assigned by guess.
register_head=515;band(audit,register_head,'Departed employees · current roster scope',1,4)
merge(audit,516,1,4,'Explicit work status only. Month does not imply a departure date. Select a name to open the master record.',MUT,PALE,10)
for j,title in enumerate(['Client / worksite','Employee','IQAMA','Work status'],1):hdr(audit,517,j,title)
jump(audit,'E515','Audit','A1','↑ Live review');jump(audit,'J3','Audit','A515','Departures →');jump(d,'B110','Audit','A515','DEPARTED STAFF  →')
audit['K4']='EMPLOYMENT EXCEPTIONS   /   Departures, reported cases and explicit source review markers; not an inference from absences.'
for rr in range(518,868):
 idx=f'$R{rr}';audit[f'R{rr}']=f'=IFERROR(MATCH(ROW()-517,{rng(M,"AQ",E)},0),"")'
 source=lambda field:f'INDEX({rng(M,field,E)},{idx})';site=source('A');prior=source('AO')
 def valid_site(ref):return 'AND('+','.join(f'{ref}<>"{txt}"' for txt in ['', '-', 'Runaway','Run away','Ran away','ZZZ INACTIVE','VACATION','Left','Resigned','Terminated'])+')'
 fields=[f'IF({valid_site(site)},{site},IF({valid_site(prior)},{prior},"Not recorded"))',f'HYPERLINK("#\'Main Dataset\'!J"&({idx}+5),{source("B")})',source('C')+'&""',source('J')]
 for j,body in enumerate(fields,1):value(audit,rr,j,f'=IF({idx}="","",{body})',fmt='@' if j==3 else None)
for st in departed_states:
 ink_rule(audit,'A518:A867',f'$D518="{st}"',BAD_TEXT);ink_rule(audit,'D518:D867',f'$D518="{st}"',BAD_TEXT)
ink_rule(audit,f'L6:L{N}','OR($M6="Terminated",$M6="Left",$M6="Resigned")',BAD_TEXT);ink_rule(audit,f'L6:L{N}','AND($K6<>"",NOT(OR($M6="Terminated",$M6="Left",$M6="Resigned")))',WARN_TEXT)
# R is used by the live rework log: hide only the lower register lookup values,
# not the physical column. No extra visible technical fields in this register.
for rr in range(518,868):audit[f'R{rr}'].number_format=';;;'
# Define the semantics beside the rules and keep navigation visible in the
# workbook. Preserve the older source's absence of dated separation history.
g['B61']='Total workforce is the in-scope master roster. Active and On Leave are separate. Inactive = Suspended + Departed; Departed = Terminated, Left or Resigned. Rankings count all in-scope payroll lines with a numeric rank, not only the first ten.'
guide(151,'Departures & client colors','Use the documented Work Status in Main Dataset. Audit → Departures lists those marked Left, Resigned or Terminated. Keep employees in the same master row. No departure date, lost pay or runaway case is inferred. Red client text = departed; amber = review / leave / suspension; green = active.')
guide(152,'Review controls','The archive retains the 73 original submissions. For new corrections, reviewer, date, evidence and a pasted revision value are checked together. A signed line is held again if the signed fields change. A resolved technical issue alone is not an authorization to pay.')
jump(g,'A153','Audit','A515','Departures  →');g['B153']='Current master status only · independent of the selected payroll period.'

# Daily-work UI: consolidate three existing supporting registers, keeping their
# data, rules, samples and source identity. No second writable copies.
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.styles import PatternFill
folds={'Clients':787,'Expense Manage':837,W:1154}
relpat=re.compile(r"'("+'|'.join(map(re.escape,folds))+r")'!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)")
def centralize(f):
 return relpat.sub(lambda m:"'Main Dataset'!"+relocated(m[2],{},folds[m[1]]),f)
# Copy the data body into the *existing* hub cells. Translate only local ranges
# in moved formulas (including absolute references); leave other source ranges.
for title,off in folds.items():
 srcws=wb[title];end={'Clients':C,'Expense Manage':X,W:C}[title];cols={'Clients':16,'Expense Manage':11,W:15}[title]
 for row in range(5,end+1):
  for j in range(1,cols+1):
   a=srcws.cell(row,j);b=master.cell(row+off,j);b.value=local_formula(a.value,{},off) if a.data_type=='f' else a.value;b._style=copy.copy(a._style)
   if row>5 and title!=W and a.data_type!='f':b.protection=unlock
 # Retain authentic samples: below the rate card/claim sections, these are
 # already excluded by the old fixed-capacity bounds. Include the yearbook.
 st=sample_specs[title][0];last=srcws.max_row;first=(48 if title==W else st)
 for row in range(first,last+1):
  for j in range(1,cols+1):
   a=srcws.cell(row,j)
   if a.value is None:continue
   b=master.cell(row+off,j);b.value=local_formula(a.value,{},off) if a.data_type=='f' else a.value;b._style=copy.copy(a._style);b.protection=xl.styles.Protection(locked=True)
   if a.comment:b.comment=copy.copy(a.comment)
   if a.hyperlink:jump(master,b.coordinate,M,'A'+str(off+5),'Back to register')
  master.row_dimensions[row+off].height=srcws.row_dimensions[row].height or 18
 for ar in srcws.merged_cells.ranges:
  if ar.min_row>=first:
   rg=CellRange(str(ar));rg.shift(row_shift=off);master.merge_cells(str(rg))
 for validation in srcws.data_validations.dataValidation:
  v=copy.copy(validation);ranges=[]
  for ar in validation.sqref.ranges:
   rg=CellRange(str(ar));rg.shift(row_shift=off);ranges.append(str(rg))
  v.sqref=' '.join(ranges);master.add_data_validation(v)
 for cf in srcws.conditional_formatting:
  rg=CellRange(str(cf.sqref));rg.shift(row_shift=off)
  for rule in srcws.conditional_formatting[cf]:
   rule=copy.deepcopy(rule)
   if rule.formula:rule.formula=[local_formula('='+f,{},off)[1:] for f in rule.formula]
   master.conditional_formatting.add(str(rg),rule)
 # Every register keeps its independent header filters; do not substitute a
 # single sheet-level filter for multiple tables on the same data hub.
 tbl=Table(displayName={'Clients':'ContractMatrix','Expense Manage':'ClaimRegister',W:'WorksiteSummary'}[title],ref=f'A{5+off}:{L(cols)}{end+off}');tbl.tableStyleInfo=TableStyleInfo(name='TableStyleMedium4',showRowStripes=False,showColumnStripes=False);master.add_table(tbl)
# Share the same controls/filters as before. The date and count summary of the
# claims register is retained directly over the consolidated source.
master['A790']='Client contracts & pay rules';master['A840']='Expense claims';master['A791'].hyperlink=None;master['A791']='Source';master['B791']='Edit the blue contract fields here. Daily and hourly rates feed timesheet checks and payroll.'
master['A841'].hyperlink=None;master['A841']='Source';master['B841']='Enter a dated claim and a positive amount. These records feed the company-wide expense report, not payroll deductions.'
merge(master,839,1,3,'Selected month · expense total',MUT,PALE,11);master['D839']=local_formula(wb['Expense Manage']['E3'].value,{},837);master['D839'].number_format=MONEY;master['E839']='Claims';master['F839']=local_formula(wb['Expense Manage']['G3'].value,{},837)
band(master,1155,'Site coverage & payroll · dashboard scope',1,15);merge(master,1156,1,15,'Roster uses the current assignment; pay uses the submitted site. Months with no submissions are not treated as absences.',MUT,PALE,11)
jump(master,'A1157',M,'A1100','Client workspace');master['D1157']='Reporting month';master['E1157']='=ViewMonth';master['E1157'].number_format=PER
for rr in range(1103,1143):
 sr=rr-1097;master[f'G{rr}']=f'=IF(A{rr}="","",HYPERLINK("#\'Main Dataset\'!A{sr+787}","Rate card"))';master[f'H{rr}']=f'=IF(A{rr}="","",HYPERLINK("#\'Main Dataset\'!A{sr+1154}","Scorecard"))'
 master[f'I{rr}']=f'=IF(A{rr}="","",HYPERLINK("#\'Main Dataset\'!"&IFERROR("C"&(MATCH(A{rr},$C$843:$C$1092,0)+842),"C842"),"Claims"))'
# Publish a short, accurate dependency guide before retiring the old surfaces.
for sh in wb:
 if sh.title in folds:continue
 for c in sh._cells.values():
  if c.data_type=='f' and isinstance(c.value,str):c.value=centralize(c.value)
  if c.hyperlink and c.hyperlink.location:
   for name,off in folds.items():
    loc=c.hyperlink.location
    if loc.startswith("'"+name+"'!"):
     ref=loc.split('!',1)[1];num=int(re.search(r'\d+',ref)[0]);c.hyperlink.location="'Main Dataset'!"+('A'+str(off+3) if num==1 else relocated(ref,{},off))
for n in wb.defined_names.values():n.attr_text=centralize(n.attr_text)
for title in folds:wb.remove(wb[title]);order.remove(title);del sample_specs[title]
removed.update(folds)
g['B22']='Main Dataset is the source for staff, contracts and expense claims. For salaries use its staff register; for the contractual daily and hourly rates use Client contracts at row 790.'
g['B53']='Seven working tabs remain. Staff, client contracts, claims and the site scorecard share Main Dataset. Time input is the sole attendance source; payroll, audit, cases and the dashboard read those records. One hidden helper supports the filters and charts.'
g['B72']='Maintain staff and contracts in Main Dataset. Paste attendance in Timesheet_Input; Salarysheet and Audit recalculate. Set policy and the compliance date here. The site scorecard and claim register are also on Main Dataset.'
g['B104']='Blank live intake: client, month, name and IQAMA in A:D; daily marks in H:AL. Payroll is calculated only from the live rows.'
g['A100']='Workbook map · seven working tabs, one hidden calculation helper'
for rr,text,addr in [(107,'Client contracts','A790'),(108,'Worksite scorecard','A1155'),(109,'Expense claims','A840')]:jump(g,'A'+str(rr),M,addr,text+'  →')
for rr in (107,108,109):g[f'B{rr}']='Section on Main Dataset, not a separate tab. '+str(g[f'B{rr}'].value)
# A team should not have to learn two nav bars. Keep only actions appropriate
# to the current sheet. The complete index remains in Control Center.
def clear_btn(sh,addr):
 c=sh[addr];c.hyperlink=None;c.value=None;c.fill=fill(BG)
for sh,keep in [(master,{'A3','H3'}),(tim,{'E3','F3'}),(pay,{'A3','C3','F3'}),(audit,{'E3','I3','J3'}),(cases,{'A3','E4'}),(g,{'A3','E3','H3'})]:
 for c in sh[3]:
  if c.hyperlink and c.coordinate not in keep:clear_btn(sh,c.coordinate)
# The master sample shortcut is in row 4, independent of the client index.
# Add one practical, scoped renewal queue in place of the older financial
# charts in the dashboard's middle. Payroll detail moves intact to Salarysheet.
for rr in range(64,105):
 for j in range(2,17):
  src=d.cell(rr,j)
  if src.value is None:continue
  dst=pay.cell(900+rr-64,j);dst._style=copy.copy(src._style)
  if src.data_type=='f':
   toks=Tokenizer(src.value)
   for tok in toks.items:
    if tok.type=='OPERAND' and tok.subtype=='RANGE' and a1.fullmatch(tok.value):tok.value="'Dashboard'!"+tok.value
   dst.value='='+''.join(z.value for z in toks.items)
  else:dst.value=src.value
  if src.hyperlink:dst.hyperlink=copy.copy(src.hyperlink);dst.hyperlink.ref=dst.coordinate
 pay.row_dimensions[rr+836].height=d.row_dimensions[rr].height or 16
for ar in list(d.merged_cells.ranges):
 if 64<=ar.min_row<=104:
  rg=CellRange(str(ar));rg.shift(row_shift=836);pay.merge_cells(str(rg));d.unmerge_cells(str(ar))
for row in range(64,106):
 for j in range(2,17):
  c=d.cell(row,j);c.value=None;c.hyperlink=None;c.fill=fill(BG)
# Re-anchor payroll visuals using columns with visible widths in the ledger.
for ch,lo,hi,j1,j2 in [(flowchart,903,916,1,8),(bc,921,935,1,8),(dc,921,935,9,16)]:
 d._charts=[original_chart for original_chart in d._charts if original_chart is not ch];ch.anchor=TwoCellAnchor(_from=AnchorMarker(col=j1,row=lo-1),to=AnchorMarker(col=j2,row=hi-1));pay.add_chart(ch)
jump(pay,'E3',P,'B900','Payroll analysis');pay['B900']='Payroll checks & rework';pay['B918']='Net payroll by site and release status';pay['B937']='Settlement detail'
# Transfer the original monetary hero calculations to this same payroll sheet.
# Keep their old addresses as nondisplayed formula drivers, not static numbers.
for col in 'BFJN':
 src=d[col+'15'];dest=pay[col+'947'];toks=Tokenizer(src.value)
 for tok in toks.items:
  if tok.type=='OPERAND' and tok.subtype=='RANGE' and a1.fullmatch(tok.value):tok.value="'Dashboard'!"+tok.value
 dest.value='='+''.join(t.value for t in toks.items);dest.number_format=src.number_format;pay[col+'946']=d[col+'14'].value;d['S'+str(6+'BFJN'.index(col))]="='Salarysheet'!"+col+'947'
# Cross-references to ready/held/gross/attendance totals follow the move.
hero_re=re.compile(r'(?<![A-Z0-9$!])([BFJN])15(?![0-9])');hero_dest={c:'S'+str(6+i) for i,c in enumerate('BFJN')}
for c in d._cells.values():
 if c.data_type=='f':c.value=hero_re.sub(lambda m:hero_dest[m[1]],c.value)
for c in pay._cells.values():
 if c.data_type=='f' and "'Dashboard'!" in c.value:
  c.value=re.sub(r"'Dashboard'!([BFJN])15\b",lambda m:"'Dashboard'!"+hero_dest[m[1]],c.value)
hide(d,'S','S')
# Show decisions, not money: expired now, upcoming renewals, missing
# submissions and holds. Retain the existing workforce totals / ranking cards.
for col,label,f,caption,fg,bg in [
 ('B','Expired Iqamas','=J108','Check expired documents',BAD_TEXT,'FDE8E6'),('F','Due in 0–15 days',"='Slicer Data'!T3",'Arrange upcoming renewals',WARN_TEXT,'FFF1CA'),('J','Missing timesheets','=F111','Active staff with no submission',WARN_TEXT,'FFF1CA'),('N','Time entries on hold','=F20','Correct the flagged source row',BAD_TEXT,'FDE8E6')]:
 for row,val,size in [(14,label,12),(15,f,24),(16,caption,11)]:
  c=d[f'{col}{row}'];c.value=val;style(c,fg,bg,size,False);c.number_format='0' if row==15 else 'General'
# A short actionable renewal list uses a single master lookup per candidate.
# The selected month only filters time records; document ages use ReviewDate.
master['AR5']='Renewal priority';master['AS5']='Renewal rank';hide(master,'AR','AS')
for rr in range(6,E+1):
 master[f'AR{rr}']=f'=IF(AND(Y{rr}=1,ISNUMBER(E{rr}),OR(E{rr}<0,E{rr}<=60)),E{rr},"")'
 master[f'AS{rr}']=f'=IF(ISNUMBER(AR{rr}),RANK(AR{rr},$AR$6:$AR${E},1)+COUNTIF($AR$6:AR{rr},AR{rr})-1,"")'
band(d,64,'04  /  Renewal worklist',2,16);merge(d,65,2,16,'First ten by expiry date. Past-due documents come first; the list includes the next 60 days. Select a name to open its master record.',MUT,WHITE,11)
segments2=[(2,5,'Employee'),(6,9,'Client'),(10,11,'Expiry date'),(12,13,'Days left'),(14,16,'Next step')]
for j1,j2,label in segments2:merge(d,66,j1,j2,label,WHITE,NAVY,11,True)
for rr in range(67,77):
 ix=f'$T{rr}';d[f'T{rr}']=f'=IFERROR(MATCH(ROW()-66,{rng(M,"AS",E)},0),"")';pick=lambda c:f'INDEX({rng(M,c,E)},{ix})'
 for (j1,j2,_),body in zip(segments2,[f'HYPERLINK("#\'Main Dataset\'!D"&({ix}+5),{pick("B")})',f'IF({pick("A")}="","Not recorded",{pick("A")})',pick('D'),pick('E'),f'IF(L{rr}<0,"Renew overdue",IF(L{rr}<=15,"Renew next","Plan renewal"))']):
  c=merge(d,rr,j1,j2,f'=IF({ix}="","",{body})',INK,WHITE if rr%2==0 else PALE,12,False);c.number_format=DATE if j1==10 else '0' if j1==12 else 'General'
 d.row_dimensions[rr].height=22
hide(d,'T','T');ink_rule(d,'L67:L76','AND(ISNUMBER(L67),L67<0)',BAD_TEXT);ink_rule(d,'N67:N76','N67="Renew overdue"',BAD_TEXT);ink_rule(d,'N67:N76','OR(N67="Renew next",N67="Plan renewal")',WARN_TEXT)
merge(d,77,2,16,f'=IF(COUNT({rng(M,"AS",E)})=0,"No dated renewals due within 60 days in this scope.","Showing "&MIN(10,COUNT({rng(M,"AS",E)}))&" of "&COUNT({rng(M,"AS",E)})&" records due by "&TEXT(N12+60,"dd mmm yyyy"))&". Missing dates: "&N29&"."',MUT,PALE,11)
# Close up the space released by moving the payroll-only detail. Preserve
# all remaining section locations, selectors, samples and chart sources.
for rr in range(78,106):d.row_dimensions[rr].hidden=True
for rr in range(135,143):d.row_dimensions[rr].hidden=True
d.print_area='A1:Q133';d.row_breaks.brk=[];d.row_breaks.append(Break(id=44));d.row_breaks.append(Break(id=77))
# Keep one row of three task buttons. Move the full link directory out of
# the daily view; users can still reach every sample and section via the guide.
for address in ('B6','F6','J6','N6'):clear_btn(d,address)
for rg in ['B6:D6','F6:H6','J6:L6','N6:P6']:d.unmerge_cells(rg)
for lo,hi,tab,loc,text in [(2,6,M,'A5','Employee records'),(7,11,T,'A6','Paste timesheet'),(12,16,'Audit','A5','Review flagged entries')]:
 c=merge(d,6,lo,hi,text,'143F55','DDECF3',12,False);c.border=Border(bottom=Side(style='thin',color='6C93A0'));c.hyperlink=Hyperlink(ref=c.coordinate,location=f"'{tab}'!{loc}",display=text)
# Remove repetitive jump strips, not measurements or working procedures.
for rr in (34,119):
 clear_btn(d,f'B{rr}');d.row_dimensions[rr].hidden=True
for address in ('N23','K7'):
 clear_btn(d,address)
d['B2']='BASMA   /   WORKFORCE OPERATIONS';d['B3']='Workforce action desk';d['B4']='="Reporting month: "&TEXT(ViewMonth,"mmmm yyyy")&"    /    "&IF(ViewSite="All","All worksites",ViewSite)&"    /    Document check: "&TEXT(N12,"dd mmm yyyy")'
d['B7']='=TEXT(\'Control Center\'!B65,"dd mmm")&" – "&TEXT(\'Control Center\'!B66,"dd mmm yyyy")&"   ·   "&\'Control Center\'!B67&" calendar days"'
d['B19']='Submitted timesheets';d['F19']='Flagged time records';d['J19']='Total workforce';d['N19']='Timesheet coverage';d['B21']='Rows in the selected month';d['F21']='Held until checks pass';d['J21']='Employees in this roster';d['N21']='Staff with a period entry'
d['B23']='Red · expired or blocked';d['F23']='Amber · due or in review';d['J23']='Green · current or ready';d['N23']='Status key';d['B24']='01  /  Submission & data quality';d['B30']='02  /  Workforce & Iqama expiry';d['B45']='03  /  Renewal horizon & work status';d['B106']='05  /  Workforce measures';d['B121']='06  /  Top ten employees';d['B122']='="Ranked by "&LOWER(ViewSort)&" in the selected month. One submitted payroll line per rank; ties retain source order."'
d['B35']='="Days from "&TEXT(N12,"dd mmm yyyy")&". Windows overlap; do not add them. Expired and undated records are excluded."'
g['B111']='Hidden calculations for the filters and the five charts (two on Dashboard; three on Salarysheet). Not a data-entry tab.'
guide(154,'Daily workflow','Start on Dashboard with the four priorities. Use the renewal worklist for employee-level follow-up. Paste timesheets only on Timesheet_Input; investigate time holds in Audit. Salarysheet retains settlement totals and charts at row 900.')
# Consolidation tests refer to these real locations instead of legacy tab names.
for j,(oldtitle,off) in enumerate(folds.items()):
 assert master.cell(5+off,1).value in ('Client / worksite','Claim date','Worksite')

# A single worklist selector keeps imminent renewals from being buried by
# a backlog of expired documents. Change view without changing employee data.
d.unmerge_cells('B65:P65');d['B65']=None
merge(d,65,2,9,'Select a name to open its expiry date in Main Dataset.',INK,WHITE,11);merge(d,65,10,12,'Show renewals',INK,PALE,11);merge(d,65,13,16,'Next 60 days',BLUE,SKY,12,True)
dv(d,'M65','list',arr(['Next 60 days','Already expired','Missing date']),blank=False,msg='Choose the renewal worklist. Other dashboard totals remain in the same scope.')
for rr in range(6,E+1):
 master[f'AR{rr}']=f'=IF(Y{rr}<>1,"",IF(\'Dashboard\'!$M$65="Missing date",IF(V{rr}="Missing date",0,""),IF(ISNUMBER(E{rr}),IF(IF(\'Dashboard\'!$M$65="Already expired",E{rr}<0,AND(E{rr}>=0,E{rr}<=60)),E{rr},""),"")))'
for rr in range(67,77):
 for col in ['J','L']:
  prev=d[f'{col}{rr}'].value;d[f'{col}{rr}']=f'=IF($M$65="Missing date","",{prev[1:]})'
 prev=d[f'N{rr}'].value;d[f'N{rr}']=f'=IF($T{rr}="","",IF($M$65="Missing date","Record expiry",{prev[1:]}))'
d['B77']=f'=IF(COUNT({rng(M,"AS",E)})=0,"No matching employees.","Showing "&MIN(10,COUNT({rng(M,"AS",E)}))&" of "&COUNT({rng(M,"AS",E)})&" employees. Earliest dates first; missing dates follow roster order.")'
ink_rule(d,'N67:N76','N67="Record expiry"',WARN_TEXT)
# Keep missing-date action labels legible without comparing blank dates.
# Shorter copy has a larger usable type size even in the merged card layout.
for addr,text in {'B8':'Report month','F8':'Worksite','J8':'Sponsor','N8':'Work status','B11':'Nationality','F11':'Contract type','J11':'Rank records by','N11':'Compliance date','B25':'Unsubmitted staff','F25':'Reporting worksites','J25':'Identity exceptions','N25':'Salaries to check','B27':'Roster IDs without an entry','F27':'Sites with submitted records','J27':'Missing or mismatched IDs / sites','N27':'Missing, text or zero base pay','B28':'Total active','F28':'Expired Iqamas · share','J28':'Iqamas over 60 days','N28':'Iqama date missing','B31':'Total inactive','F31':'Total ranked records','J31':'Iqama · 9–15 days','N31':'Iqama · 1–14 days','B33':'Suspended or departed','F33':'All records in the ranking','J33':'Includes days 9 and 15','N33':'Includes days 1 and 14','B47':'Employees by non-overlapping expiry range','J47':'Employees by current work status','B107':'Active workforce share','F107':'On leave','J107':'Expired Iqamas','N107':'Inside Iqama alert window','F110':'Active without timesheet','J110':'Both documents current','N110':'Iqama · 31–60 days','B112':'Present equivalents','F112':'Recorded absences','J112':'Recorded days off','N112':'Verified record share','B116':'Insurance expired','F116':'Insurance due soon','J116':'Full-time staff','N116':'Average ready pay · SAR'}.items():d[addr]=text
for r,limits in [(37,[60,30,20,15]),(41,[14,9,1,0])]:
 for col,n in zip(starts,limits):d.cell(r,col).value='Due today' if n==0 else 'Within '+str(n)+(' day' if n==1 else ' days')
for address in ('B110','B119','B136','F136','J136','N136','B138','F138','J138','N138','B139','J139','B140','B142'):
 ce=d[address]
 if ce.hyperlink:ce.hyperlink=None
# Use an actual dashboard reference for the scope of the concentration cards.
d['B110']='Departed staff';d['B110'].comment=Comment('Left, resigned or terminated. The list is in Audit → Departures.','Basma / Operations')
# Source and archive sections are linked from a single directory; remove
# residual secondary toolbar strips from hidden dashboard navigation.
for rr in (136,138,139,140,142):
 for j in (2,6,10,14):
  ce=d.cell(rr,j)
  if ce.value is not None:clear_btn(d,ce.coordinate)
for sh,links in [(master,[('A3','Dashboard'),('H3','Client workspace'),('A4','Sample row')]),(tim,[('E3','Filter entries'),('F3','Sample row')]),(pay,[('A3','Dashboard'),('C3','Sample row'),('E3','Payroll analysis'),('F3','Client statement')]),(audit,[('E3','Rework log'),('I3','Sample row'),('J3','Departures')]),(cases,[('A3','Edit case fields')]),(g,[('A3','Dashboard'),('E3','Workbook map'),('H3','Sample rows')])]:
 for addr,text in links:
  ce=sh[addr]
  if ce.hyperlink:ce.value=text+'  ›';ce.hyperlink.display=ce.value
# Guide smaller controls into their new single source without claiming that
# old external sheets still exist. Archive/source quotations are left intact.
for rr in [22,23,33,34,35,36,37,38,39,40,41,45,59,72,73,74,78,79,80,81,82,83,84,85,86,87,88,89,90,91,92,93,94,95,96,97]:
 c=g[f'B{rr}']
 if c.data_type=='s' and c.value:c.value=c.value.replace('in Clients','in Main Dataset → Client contracts').replace('on Clients','in Main Dataset → Client contracts')
# Preserve ability to reach section samples after retiring the old tab bars.
for title,top in [('Clients',47),('Expense Manage',257),(W,64)]:
 off=folds[title];row=next(r for r in range(142,150) if g[f'A{r}'].hyperlink and g[f'A{r}'].hyperlink.location==f"'{M}'!A{top+off}")
 g[f'A{row}'].value={'Clients':'Rate card example',W:'Site scorecard example','Expense Manage':'Expense example'}[title]+'  ›'
# No additional source data. The dynamic renewal rows are navigation, not
# tickets or confirmations that a renewal has been completed.
guide(155,'Renewal worklist','Choose Next 60 days, Already expired or Missing date on Dashboard. The first ten matching employees follow all master filters and the compliance date. Select a name to verify or update the record. Blank dates are never interpreted as overdue.')

def apply_readable_ui():
 """Final style pass, after measured widths. No logic changes here."""
 # Keep the same grid and section landmarks. Dark text on light backgrounds,
 # white on navy/teal. No pale-grey captions or 9pt controls in daily use.
 for sh in wb:
  if sh.sheet_state!='visible':continue
  sh.sheet_view.zoomScale=100
  # Only cells that actually display content need a new font, so a large
  # prewired empty range is not expanded or given fake placeholder data.
  for ce in sh._cells.values():
   if ce.value is None or ce.font.color is None:continue
   ft=ce.font;fcolor=ft.color.rgb[-6:] if ft.color.type=='rgb' else INK;size=ft.sz or 10
   if size<11 or fcolor in [MUT,BLUE,AQUA]:
    color={MUT:'30474B',BLUE:'184968',AQUA:'D6FFF0'}.get(fcolor,fcolor);ce.font=font(color,max(11,size),bool(ft.b))
  if sh is not d:
   sh.sheet_format.defaultRowHeight=18;sh.row_dimensions[3].height=max(22,sh.row_dimensions[3].height or 0)
   if sh is not g:sh.row_dimensions[5].height=31
   for rr in (1,2):
    ce=sh.cell(rr,1);ce.font=font(WHITE if rr==1 else '243A40',18 if rr==1 else 11,False);ce.alignment=align('left',1,True)
   sh.row_dimensions[1].height=30;sh.row_dimensions[2].height=32
 for j in range(1,18):d.column_dimensions[L(j)].width=2.3 if j in (1,5,9,13,17) else 12
 d.sheet_view.zoomScale=100
 d['B2'].font=font('D6FFF0',11,False);d['B3'].font=font(WHITE,23,False);d['B4'].font=font(WHITE,12,False);d.row_dimensions[2].height=20;d.row_dimensions[3].height=34;d.row_dimensions[4].height=24;d.row_dimensions[6].height=27
 for rr in [8,11]:
  d.row_dimensions[rr].height=19;d.row_dimensions[rr+1].height=28
  for col in starts:
   h=d.cell(rr,col);h.font=font('213B40',11,False);v=d.cell(rr+1,col);v.font=font('173D52',13,False)
 for rr in (14,19,25,28,31,37,41,107,110,112,116):
  for col in starts:
   h=d.cell(rr,col);v=d.cell(rr+1,col);h.font=font('263D41',12,False);h.alignment=align('left',1,True)
   co=v.font.color.rgb[-6:] if v.font.color and v.font.color.type=='rgb' else INK;v.font=font(co,22,False)
  d.row_dimensions[rr].height=25;d.row_dimensions[rr+1].height=32
 for rr in [16,21,27,33,39,43,109,114,118]:
  d.row_dimensions[rr].height=29
  for col in starts:d.cell(rr,col).font=font('34494C',11,False);d.cell(rr,col).alignment=align('left',1,True)
 for rr in [24,30,45,64,106,121]:
  c=d.cell(rr,2);style(c,WHITE,NAVY,14,False);d.row_dimensions[rr].height=28
 for rr in (35,62,65,77,122):d.row_dimensions[rr].height=32
 for ce in ('B35','B62','J62','B65','B77','B122'):d[ce].alignment=align('left',1,True)
 # Show a true count of zero as 0; reserve the dash for undefined ratios.
 for row in [15,20,26,29,32,38,42,108,111,113,117]:
  for col in starts:
   cell=d.cell(row,col)
   if cell.number_format==INT:cell.number_format='#,##0'
 # Never show empty pseudo-tasks when a selected month has no time submissions.
 d['B23'].alignment=align('left',1,False)
 d['B35'].alignment=align('left',1,True)
 for rr in range(67,77):
  for col in (2,6):d.cell(rr,col).alignment=Alignment(horizontal='left',vertical='center',indent=1,shrinkToFit=True)
 d.row_dimensions[62].height=43
 # On the master ledger, calculation helpers in the relocated community of
 # tables are needed by formulas, not for reading down the table.
 for rr in range(793,833):master.cell(rr,16).number_format=';;;'
 for rr in range(843,1093):master.cell(rr,11).number_format=';;;'
 for rr in range(1160,1200):master.cell(rr,15).number_format=';;;'
 for r0,j0 in [(792,16),(842,11),(1159,15)]:master.cell(r0,j0).comment=Comment('Calculation helper; not an input.','Basma / Operations')
 for ss,ad in [(master,'A1100'),(master,'A790'),(master,'A840'),(master,'A1155'),(pay,'B900'),(pay,'B918'),(pay,'B937')]:
  c=ss[ad];style(c,WHITE,NAVY,14,False);ss.row_dimensions[c.row].height=28

# Presentation revision: a concise monthly report, never a second input source.
# Day evidence is still necessary for attendance/OT rules, but it need not be
# reproduced on screen or in a client-facing monthly summary.
monthly_columns={1:('Client / worksite','A'),2:('Month / year','B'),3:('Employee','C'),4:('IQAMA','D'),5:('Days logged','AP'),6:('Check result','BQ'),7:('Flag / action','BR'),8:('Remarks','AU'),9:('Source row',None)}
for j,(label,_) in monthly_columns.items():tim.cell(434,j).value=label
for r in range(435,855):
 ix=f'$BU{r}'
 for j,(_,col) in monthly_columns.items():
  ref=f'INDEX(${col}$6:${col}${N},{ix})' if col else None
  f=f'IF(ISBLANK({ref}),"",{ref})' if ref else f'HYPERLINK("#\'Timesheet_Input\'!A"&({ix}+5),"Open row "&({ix}+5))'
  c=tim.cell(r,j);c.value=f'=IF({ix}="","",{f})';c.number_format=PER if j==2 else '0.##' if j==5 else '@' if j==4 else 'General';c.alignment=align('left',1,False)
value(tim,432,8,'Month scope');value(tim,432,9,'All periods',True).protection=unlock;dv(tim,'I432','list',arr(['All periods','Dashboard month']),blank=False,msg='All periods includes invalid or missing months. Dashboard month restricts the list to the selected reporting month.')
for r in range(6,N+1):tim[f'BT{r}']=f'=IF(AND(BN{r}=1,OR($B$432="All",A{r}=$B$432),OR($D$432="All",BQ{r}=$D$432),OR($I$432="All periods",B{r}=ViewMonth)),COUNT($BT$5:BT{r-1})+1,"")'
ink_rule(tim,'A435:A854','$F435="Ready"',OK_TEXT);ink_rule(tim,'A435:A854','$F435="Hold"',BAD_TEXT)
nm('MonthlyTimesheetReview',f"'{T}'!$A$430:INDEX('{T}'!$A$430:$I$854,5+'{T}'!$F$432,9)")
tim['A430']='Monthly timesheet review';tim['A431']='Filter by client, status and month scope. Each line shows the check and any source remark. Open its source row to make corrections.'
jump(tim,'E3',T,'A430','Monthly review ▾');tim['E3'].comment=Comment('Filter the monthly review without changing the live input, the payroll ledger or the original row numbers.','Basma / Operations')
tim['A1']='Monthly timesheets';tim['A2']='One employee per month. Paste values in the blue input blocks: client, month start, name and IQAMA in A:D; daily marks in H:AL. Checks run automatically.'
tim['B4']='1  ·  Employee and period';tim['E4']='2  ·  Matched to the master';tim['H4']='3  ·  Attendance evidence   /   P, A, OFF, 0.5P or hours'
# Consistent monthly identifiers across input, payroll, client statement and
# audit queue. Only the fields on the affected rows move. Preserve archived
# evidence, signatures, formulas, samples and editable areas as before.
field_moves={P:({1:2,2:1},[(6,N),(428,428),(436,855)]),'Audit':({1:5,5:1},[(6,N),(509,509)])}
def moved_field(ref,cols,intervals):
 parts=ref.split(':');ms=[a1.fullmatch(t) for t in parts]
 if not all(ms) or not all(any(lo<=int(m[4])<=hi for lo,hi in intervals) for m in ms):return ref
 if len(ms)==2 and ms[0][2]!=ms[1][2]:return ref
 return relocated(ref,cols)
for sn,(mapping,zones) in field_moves.items():
 sheet=wb[sn];keys=list(mapping);a,b=keys;rows=([5,427,435] if sn==P else [5,508])+[r for lo,hi in zones for r in range(lo,hi+1)]
 for r in rows:
  x,y=sheet.cell(r,a),sheet.cell(r,b);x.value,y.value=y.value,x.value;x._style,y._style=copy.copy(y._style),copy.copy(x._style)
 for c in sheet._cells.values():
  if c.data_type!='f' or not isinstance(c.value,str) or not re.search(r'(?<![\w!])\$?['+''.join(L(k) for k in keys)+r']\$?\d+',c.value):continue
  toks=Tokenizer(c.value)
  for tok in toks.items:
   if tok.type=='OPERAND' and tok.subtype=='RANGE' and '!' not in tok.value:tok.value=moved_field(tok.value,mapping,zones)
  c.value='='+''.join(tok.value for tok in toks.items)
 rp=re.compile("'"+re.escape(sn)+r"'!(\$?["+''.join(L(k) for k in keys)+r"]\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)\b")
 for ws in wb:
  for c in ws._cells.values():
   if c.data_type=='f' and "'"+sn+"'!" in c.value:c.value=rp.sub(lambda mt:"'"+sn+"'!"+moved_field(mt[1],mapping,zones),c.value)
 for nmdef in wb.defined_names.values():nmdef.attr_text=rp.sub(lambda mt:"'"+sn+"'!"+moved_field(mt[1],mapping,zones),nmdef.attr_text)
for rr in (5,427,435):pay[f'A{rr}']='Client / worksite';pay[f'B{rr}']='Month / year'
for rr in (5,508):audit[f'A{rr}']='Submitted client';audit[f'B{rr}']='Month / year';audit[f'E{rr}']='Time row'
for rr in range(6,N+1):
 audit[f'A{rr}'].number_format='General';audit[f'B{rr}'].number_format=PER
# Ensure the revised sample has a visible, real display format (rather than
# a format inherited from the source-row hyperlink). The row stays outside pay.
pay['A428'].number_format='General';pay['B428'].number_format=PER
for r in range(6,N+1):pay[f'A{r}'].number_format='General';pay[f'B{r}'].number_format=PER
for r in range(436,856):pay[f'A{r}'].number_format='General';pay[f'B{r}'].number_format=PER
for tab,cols in [(pay,('A','B')),(audit,('A','E'))]:tab.column_dimensions['A'].width=25
pay['A2']='One employee per client and month. Ready net + held net = estimated net. View the client statement for a concise, filterable monthly report.'
# Show the documented case remark, if one exists. Do not infer a case from
# an absence, insert allegations into a general staff note or duplicate cases.
master['AT5']='Case remark';master.column_dimensions['AT'].width=36
for r in range(6,E+1):value(master,r,46,None,True).protection=unlock
master['AT5'].comment=Comment('Optional case-specific note based on supporting evidence. The linked register shows this note only when the case is explicitly recorded.','Basma / Operations')
cases['G5']='Case remark';cases.column_dimensions['G'].width=38;cases.auto_filter.ref=f'A5:G{E}';cases.print_area=f'A1:G{E}'
for r in range(6,E+1):
 ref=f'INDEX({rng(M,"AT",E)},$F{r})';value(cases,r,7,f'=IF($F{r}="","",IF(ISBLANK({ref}),"",{ref}))')
for r in (1,2):
 cases.unmerge_cells(f'A{r}:E{r}');cases.merge_cells(f'A{r}:G{r}')
cases['G358']='Case remark';value(cases,359,7,'Awaiting source confirmation.');sample_specs[case_tab]=(357,'E4',7,'Fictional case · not a reported employee');cases['F358']='Source index'
cases['A2']='Only explicitly reported cases appear here. Select an employee to update the source record. Resolved cases stay in the register.'
g['B116']='The master case fields are status, reported month and client in AM:AO, plus an optional documented case remark in AT. Red client names = reported, amber = under review, green = resolved. No time or pay rule creates a case.'
# Keep the full original October examples for auditability, but show one row.
g.row_dimensions.group(127,138,outline_level=1,hidden=True);g.row_dimensions[139].collapsed=True;g['A121']='One original example row is visible. Expand the grouped rows to inspect the remaining source archive; no examples enter payroll.'
g['B104']='Timesheet_Input is the only attendance source. Paste into A:D and H:AL; Monthly review shows one client/employee/month row with the result, flag and remark. Filter by client, release state or month scope.'
# De-emphasize duplicate captions rather than deleting any metrics; the useful
# charts and action list move into view sooner. Ordinary Excel outline toggles
# restore the supplementary descriptions without changing the worksheet data.
for rr in (21,27,33,39,43):
 d.row_dimensions.group(rr,rr,outline_level=1,hidden=True);d.row_dimensions[rr+1].collapsed=True
for rr in (25,28,31,37,41):
 for j in starts:
  text=d.cell(rr,j).value or '';ce=d.cell(rr+1,j)
  if ce.comment is None:ce.comment=Comment(str(text)+' • '+str(d.cell(rr+2,j).value or 'Filtered workforce measure.'),'Basma / Operations')
# A single plain-language decision line summarises the native charts/cards.
merge(d,17,2,16,'=IF(J20=0,"No employees match these filters.","Renewal workload: "&TEXT(B15,"#,##0")&" overdue  ·  "&TEXT(F15,"#,##0")&" due in 0–15 days  ·  "&TEXT(N29,"#,##0")&" missing expiry dates.")','243A40','EBF1F4',11,False);d.row_dimensions[17].height=22
# Distinct time-entry visual language: slate-blue identity band, white input
# paper, dark text; the actual positive/negative/hold text rules stay intact.
def monthly_finish():
 blue='24465F';slate='20313B';light='E9F2F7';tint='F7FAFC'
 tim.sheet_properties.tabColor=blue;style(tim['A1'],WHITE,blue,18,False);style(tim['A2'],slate,light,11,False,'left',1,True);tim.row_dimensions[2].height=34
 for address,fg,bg in [('A4',blue,light),('B4',blue,light),('E4',slate,'EFF2EF'),('H4',blue,light)]:style(tim[address],fg,bg,11,False)
 for j in range(1,72):
  c=tim.cell(5,j)
  if c.value is None:continue
  editable=j in [1,2,3,4,*range(8,39),44,45,46,47,50];style(c,WHITE,blue if editable else NAVY,10,True,'center' if 8<=j<=38 else 'left',0 if 8<=j<=38 else 1,True)
 for rr in range(6,N+1):
  for j in [1,2,3,4,*range(8,39),44,45,46,47]:
   c=tim.cell(rr,j);c.fill=fill(WHITE if rr%2==0 else tint);c.font=font('193D55',11,False)
   if j in (8,15,22,29,36):c.border=Border(left=Side(style='thin',color='AFBFCA'))
 for j in (3,7,8):
  for rr in range(435,855):tim.cell(rr,j).alignment=Alignment(horizontal='left',vertical='center',indent=1,shrinkToFit=True)
 tim['A430'].font=font(WHITE,16,False);tim['A430'].fill=fill(blue);tim.row_dimensions[430].height=27;tim.row_dimensions[431].height=29
 for j,wi in [(8,20),(9,16)]:tim.column_dimensions[L(j)].width=max(wi,tim.column_dimensions[L(j)].width)
 for a in ['B901','J901','B919','J919']:
  c=pay[a];c.font=font(INK,14,False)
 for r in [46,65]:
  for col in [2,10]:
   c=d.cell(r,col)
   if c.value is not None and c.data_type!='f':c.font=font(INK,13 if r==46 else 11,False)
 # Only the table headers are bold. Titles and stage labels use normal weight.
 for c in ['B24','B30','B45','B64','B106','B121']:d[c].font=font(WHITE,14,False)
 for sh in (pay,audit,master,cases):
  sh.sheet_view.showGridLines=False;sh.row_dimensions[1].height=30
 for j in [1,2,3,4,5,7]:style(cases.cell(5,j),WHITE,TEAL,10,True,'left',1,True)
 cases.row_dimensions[2].height=24

# Database handoff: formalize existing operational ranges rather than copy
# them to another workbook. Supporting reports remain formula-backed views.
# The case register shares the history tab; the new journal retains dated
# facts, not reconstructed transitions or a falsely automatic edit history.
old_case_tab=case_tab;case_tab='Status History';old_case_sheet=cases
hist=wb.create_sheet(case_tab,wb.index(old_case_sheet));setup(hist);view_lock(hist);hist.sheet_properties.tabColor=TEAL
hist_end=425;case_offset=430;case_start=6+case_offset;case_end=E+case_offset
# Relocate the complete case projection and its reference example, including
# source links, status rules and evidence. Master fields are not duplicated.
for ce in old_case_sheet._cells.values():
 if ce.value is None:continue
 dst=hist.cell(ce.row+case_offset,ce.column);dst._style=copy.copy(ce._style)
 dst.value=local_formula(ce.value,{},case_offset).replace('ROW()-5,','ROW()-435,') if ce.data_type=='f' else ce.value
 if ce.comment:dst.comment=copy.copy(ce.comment)
 if ce.hyperlink:
  lk=copy.copy(ce.hyperlink);lk.ref=dst.coordinate;dst.hyperlink=lk
for area in old_case_sheet.merged_cells.ranges:
 rg=CellRange(str(area));rg.shift(row_shift=case_offset);hist.merge_cells(str(rg))
for rr,dim in old_case_sheet.row_dimensions.items():hist.row_dimensions[rr+case_offset].height=dim.height
for cf in old_case_sheet.conditional_formatting:
 rg=CellRange(str(cf.sqref));rg.shift(row_shift=case_offset)
 for item in old_case_sheet.conditional_formatting[cf]:
  rule=copy.deepcopy(item)
  if rule.formula:rule.formula=[local_formula('='+f,{},case_offset)[1:] for f in rule.formula]
  hist.conditional_formatting.add(str(rg),rule)
# Only link targets change; no historical record refers to an obsolete tab.
for sh in wb:
 if sh==old_case_sheet:continue
 for ce in sh._cells.values():
  if ce.hyperlink and ce.hyperlink.location and ce.hyperlink.location.startswith("'"+old_case_tab+"'!"):
   lk=ce.hyperlink;lk.location=f"'{case_tab}'!"+relocated(lk.location.rsplit('!',1)[1],{},case_offset)
wb.remove(old_case_sheet);cases=hist;order=wb.sheetnames;removed.add(old_case_tab)
del sample_specs[old_case_tab];sample_specs[case_tab]=(426,'B3',10,'Example movement · not a live record or a master update')
# The 70 actual source dates are joining dates only. Clients, prior statuses
# and undocumented departure/case dates are intentionally not inferred.
join_events=[(i+3,cleanid(v[2]),datev(v[10])) for i,v in enumerate(ep) if isinstance(datev(v[10]),dt.datetime)]
merge(hist,1,1,10,'Status & movement history',WHITE,NAVY,18,False);merge(hist,2,1,10,'One event per row. Record the effective date and evidence before updating the current master. Earlier events stay unchanged.',INK,PALE,11,False)
hist.row_dimensions[1].height=30;hist.row_dimensions[2].height=25;hist.sheet_format.defaultRowHeight=18
jump(hist,'A3',M,'A5','Employee master');jump(hist,'B3',case_tab,'A426','Sample row');jump(hist,'C3',case_tab,'A431','Reported cases');value(hist,3,6,'Flagged events');value(hist,3,7,f'=COUNT(K6:K{hist_end})',fmt='0');hist.row_dimensions[3].height=23
merge(hist,4,1,10,'Append verified movements in the next empty row. Record corrections as new entries. Joining records from the source have no assumed client or previous status.',MUT,WHITE,11);hist.row_dimensions[4].height=25
hist_heads=['Client at event','Effective date','IQAMA','Employee · master','Event ID','Event type','From value','To value','Evidence / reference','Record check','Exception order'];types=['Joining','Status change','Client transfer','Case update','Exit','Correction']
for j,(label,ww) in enumerate(zip(hist_heads,[23,17,18,35,23,20,23,23,45,28,16]),1):
 hdr(hist,5,j,label);hist.column_dimensions[L(j)].width=ww
hist.row_dimensions[5].height=25;hist.freeze_panes='D6';hist.print_title_rows='1:5';hist.print_area=f'A1:J{5+len(join_events)}'
def literal_criteria(ref):return f'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE({ref},"~","~~"),"*","~*"),"?","~?")'
for rr in range(6,hist_end+1):
 seed=join_events[rr-6] if rr-6<len(join_events) else None
 data={2:seed[2],3:seed[1],5:'JOIN-'+seed[1],6:'Joining',9:f'Import: Employees!K{seed[0]}'} if seed else {}
 for j in [1,2,3,5,6,7,8,9]:
  ce=value(hist,rr,j,data.get(j),not bool(seed),DATE if j==2 else '@' if j in [3,5] else None);ce.protection=xl.styles.Protection(locked=bool(seed))
 value(hist,rr,4,f'=IF(C{rr}="","",IFERROR(INDEX({rng(M,"B",E)},MATCH(TRIM(C{rr}&""),{rng(M,"Z",E)},0)),"Not in master"))')
 exact_id=f'IFERROR(AND(LEN(C{rr}&"")=10,TEXT(VALUE(C{rr}),"0000000000")=C{rr}&""),FALSE)';valid_date=f'IFERROR(AND(ISNUMBER(B{rr}),B{rr}=INT(B{rr}),B{rr}>0,B{rr}<=TODAY()),FALSE)'
 oneof=lambda col,options:'OR('+','.join(f'{col}{rr}="{op}"' for op in options)+')'
 gcrit=literal_criteria(f'G{rr}&""');hcrit=literal_criteria(f'H{rr}&""')
 checks=[(f'LEN(TRIM(E{rr}&""))=0','Event ID missing'),(f'COUNTIF($E$6:$E${hist_end},{literal_criteria("E"+str(rr))})>1','Duplicate event ID'),(f'NOT({exact_id})','Invalid IQAMA'),(f'COUNTIF({rng(M,"Z",E)},C{rr})<>1','Check employee ID'),(f'NOT({valid_date})','Check effective date'),(f'NOT({oneof("F",types)})','Check event type'),(f'LEN(TRIM(I{rr}&""))=0','Evidence missing'),(f'AND(F{rr}="Client transfer",OR(G{rr}="",H{rr}="",TRIM(G{rr})=TRIM(H{rr})))','Check transfer sites'),(f'AND(F{rr}="Status change",OR(NOT({oneof("G",work_states)}),NOT({oneof("H",work_states)}),G{rr}=H{rr}))','Check status change'),(f'AND(F{rr}="Exit",NOT({oneof("H",departed_states)}))','Check exit status'),(f'AND(F{rr}="Case update",NOT({oneof("H",case_states)}))','Check case status'),(f'AND(F{rr}="Correction",OR(G{rr}="",H{rr}="",G{rr}=H{rr}))','Check correction'),(f'COUNTIFS($B$6:$B${hist_end},B{rr},$C$6:$C${hist_end},C{rr},$F$6:$F${hist_end},F{rr},$G$6:$G${hist_end},{gcrit},$H$6:$H${hist_end},{hcrit})>1','Duplicate event')]
 body='"Recorded"'
 for test,label in reversed(checks):body=f'IF({test},"{label}",{body})'
 value(hist,rr,10,f'=IF(COUNTA(A{rr}:C{rr},E{rr}:I{rr})=0,"",{body})')
 value(hist,rr,11,f'=IF(AND(J{rr}<>"",J{rr}<>"Recorded",OR(\'Audit\'!$D$3="All periods",IFERROR(AND(YEAR(B{rr})=YEAR(ViewMonth),MONTH(B{rr})=MONTH(ViewMonth),OR(ViewSite="All",A{rr}=ViewSite)),FALSE))),COUNT($K$5:K{rr-1})+1,"")')
 # The underlying journal's date is separate from the lower live case list.
for rng0,typ,one,two,op,msg in [(f'B6:B{hist_end}','date','DATE(1900,1,1)','TODAY()','between','Use the actual documented event date, not the payroll month.'),(f'F6:F{hist_end}','list',arr(types),None,None,'Choose a movement type.'),(f'C6:C{hist_end}','custom','IFERROR(AND(LEN(C6&"")=10,TEXT(VALUE(C6),"0000000000")=C6&""),FALSE)',None,None,'Use a 10-digit IQAMA present in EmployeeMaster.')]:dv(hist,rng0,typ,one,two,op,msg)
hide(hist,'K','K');ink_rule(hist,f'J6:J{hist_end}','AND(J6<>"",J6<>"Recorded")',BAD_TEXT);ink_rule(hist,f'J6:J{hist_end}','J6="Recorded"',OK_TEXT)
merge(hist,426,1,8,'Sample row  /  Client transfer · illustration only; not part of the history table',MUT,PALE,11,False);jump(hist,'J426',case_tab,'A5','Back to history')
for j,field in enumerate(hist_heads[:10],1):hdr(hist,427,j,field)
for j,v in enumerate(['Example client',start,'0000000000','Example employee','EXAMPLE-001','Client transfer','Example site A','Example site B','Documented transfer order','Illustration only'],1):value(hist,428,j,v,fmt=DATE if j==2 else '@' if j==3 else None)
hist['A431']='Reported runaway cases';hist['A432']='Current cases from the master, including resolved records. These are not dated movement events. Unknown dates and clients remain unassigned.'
# Ensure no blank technical index is exposed in the case view, but leave F
# visible because it is the useful event type in the journal above it.
for rr in range(case_start,case_end+1):hist[f'F{rr}'].number_format=';;;'
hist['F435']='';hist['F788']='';jump(hist,'A433',M,'AM5','Edit case fields');jump(hist,'B433',case_tab,'A5','Back to history');jump(hist,'E434',case_tab,'A787','Case example')
# Centralize exceptions without disturbing archived evidence, signed reviews or
# amounts. The existing top queue displays pay holds, then master issues, then
# history flags. Pages cover every prepared record; no exception is truncated.
master['AU5']='Database check';master['AV5']='Master exception order';hide(master,'AV','AV');master.column_dimensions['AU'].width=23
for rr in range(6,E+1):
 ident=f'IFERROR(AND(LEN(Z{rr})=10,TEXT(VALUE(Z{rr}),"0000000000")=Z{rr}),FALSE)';valid_status='OR('+','.join(f'J{rr}="{st}"' for st in work_states)+')'
 checks=[(f'B{rr}=""','Employee name missing'),(f'NOT({ident})','Invalid master IQAMA'),(f'COUNTIF($Z$6:$Z${E},Z{rr})<>1','Duplicate master ID'),(f'NOT({valid_status})','Unrecognized status'),(f'A{rr}=""','Client missing'),(f'X{rr}="Check salary"','Check master salary'),(f'V{rr}="Missing date"','Missing IQAMA date'),(f'W{rr}="Missing date"','Missing insurance date')];body='"OK"'
 for test,msg in reversed(checks):body=f'IF({test},"{msg}",{body})'
 value(master,rr,47,f'=IF(COUNTA(A{rr}:D{rr},F{rr}:U{rr},AM{rr}:AO{rr},AT{rr})=0,"",{body})')
 value(master,rr,48,f'=IF(AND(AU{rr}<>"",AU{rr}<>"OK",OR(\'Audit\'!$D$3="All periods",Y{rr}=1)),COUNT($AV$5:AV{rr-1})+1,"")')
ink_rule(master,f'AU6:AU{E}','AND(AU6<>"",AU6<>"OK")',BAD_TEXT)
# Include high-risk config inputs in the same queue. A broken pay policy must
# not release payroll: use safe arithmetic fallbacks in the amount estimate,
# and block release until invalid settings are repaired.
config_checks={7:'AND(ISNUMBER(B7),B7=INT(B7),B7>=0,B7<=365)',8:'AND(ISNUMBER(B8),B8=INT(B8),B8>=0,B8<=365)',9:'AND(ISNUMBER(B9),B9>=0)',10:'AND(ISNUMBER(B10),B10=INT(B10),B10>=DATE(2000,1,1),B10<=DATE(2100,12,31))',11:'AND(ISNUMBER(B11),B11=INT(B11),B11>=1,B11<=31)',12:'AND(ISNUMBER(B12),B12=INT(B12),B12>=1,B12<=31)',13:'AND(ISNUMBER(B13),B13>=0,B13<=1)',14:'AND(ISNUMBER(B14),B14>=0)',15:'AND(ISNUMBER(B15),B15>=0)',16:'AND(ISNUMBER(B16),B16>=0,B16<=1)',17:'AND(ISNUMBER(B17),B17>=0)',18:'B18="SAR"'}
for rr in range(6,19):g.unmerge_cells(f'C{rr}:H{rr}')
g['C6']='Purpose';g['D6']='Rule check';g['E6']='Exception order';g.column_dimensions['C'].width=53
for rr,test in config_checks.items():
 g[f'C{rr}'].alignment=align('left',1,True);g.row_dimensions[rr].height=33;value(g,rr,4,f'=IF(IFERROR({test},FALSE),"OK","Check rule")');value(g,rr,5,f'=IF(D{rr}="Check rule",COUNT($E$6:E{rr-1})+1,"")');g[f'E{rr}'].number_format=';;;'
ink_rule(g,'D7:D18','D7="Check rule"',BAD_TEXT)
for rr in range(6,N+1):
 # Production arithmetic uses these settings directly: a bad string in a
 # pasted policy must result in a hold, rather than a #VALUE! exception.
 for col in ['G','I']:
  ce=pay[f'{col}{rr}'];f=ce.value
  for rule_row in [11,12,13]:f=f.replace(f"'Control Center'!$B${rule_row}",f"N('Control Center'!$B${rule_row})")
  ce.value=f
 base=pay[f'AJ{rr}'].value[1:];pay[f'AJ{rr}']=f'=IF(\'{T}\'!BN{rr}=0,"",IF(OR(COUNTIF(\'Control Center\'!$D$11:$D$13,"Check rule")>0,\'Control Center\'!$D$18<>"OK"),"Check payroll settings",{base}))'
# Four compact counts feed the entire exception queue. The reference date for
# history scope uses effective month + client, independently of payroll checks.
summary=ls
for addr,formula in [('AB2',f'=COUNT({rng(P,"X")})'),('AB3',f'=COUNT({rng(M,"AV",E)})'),('AB4',f'=COUNT({rng(case_tab,"K",hist_end)})'),('AB5',"=COUNT('Control Center'!$E$7:$E$18)")]:summary[addr]=formula
ls['AA1']='Audit queue counts';ls['AB1']='Records';ls['AA2']='Payroll';ls['AA3']='Employee master';ls['AA4']='Movement history';ls['AA5']='Configuration'
value(audit,3,6,'Page');value(audit,3,7,1,True,'0').protection=unlock;dv(audit,'G3','whole',1,4,'between','420 findings per page; choose 1–4.');audit.column_dimensions['G'].width=28
merge(audit,4,18,25,'Time-entry rework · fixed row, evidence and reviewer sign-off',WHITE,TEAL,11,False)
audit['A1']='Exceptions & audit';audit['A2']='Payroll holds, master-data errors, movement flags and invalid rules. Fix at the source. Choose All periods to include undated or unmatched records.'
audit['A4']='Open issues';audit['B4']='=SUM(\'Slicer Data\'!AB2:AB5)&" findings  /  "&MAX(1,ROUNDUP(SUM(\'Slicer Data\'!AB2:AB5)/420,0))&" page(s). One primary issue per record; recheck after each correction."'
for j,txt in enumerate(['Client','Month / year','Employee','IQAMA','Source record','First review item','Expected name','Expected site','Held net · SAR','Review action'],1):audit.cell(5,j).value=txt
for j,txt in [(27,'Queue number'),(28,'Issue source'),(29,'Source index')]:ls.cell(460,j,txt)
for j in range(1,10):audit.cell(508,j).value=audit.cell(5,j).value
for r in range(6,N+1):
 h=r+455;num=f"'Slicer Data'!AA{h}";kind=f"'Slicer Data'!AB{h}";idx=f"'Slicer Data'!AC{h}";hcount="MAX(0,MIN(3,IFERROR(INT('Audit'!$G$3)-1,0)))"
 ls[f'AA{h}']=f'=ROW()-460+420*{hcount}';ls[f'AB{h}']=f'=IF(AA{h}<=$AB$2,"Time",IF(AA{h}<=SUM($AB$2:$AB$3),"Employee",IF(AA{h}<=SUM($AB$2:$AB$4),"History",IF(AA{h}<=SUM($AB$2:$AB$5),"Rules",""))))'
 ls[f'AC{h}']=f'=IFERROR(IF(AB{h}="Time",MATCH(AA{h},{rng(P,"X")},0),IF(AB{h}="Employee",MATCH(AA{h}-$AB$2,{rng(M,"AV",E)},0),IF(AB{h}="History",MATCH(AA{h}-SUM($AB$2:$AB$3),{rng(case_tab,"K",hist_end)},0),IF(AB{h}="Rules",MATCH(AA{h}-SUM($AB$2:$AB$4),\'Control Center\'!$E$7:$E$18,0),0)))),0)'
 def fld(sh,col,end):
  ref=f'INDEX({rng(sh,col,end)},{idx})';return f'IF(ISBLANK({ref}),"",{ref})'
 for j,pcol,mcol,hcol in [(1,'A','A','A'),(2,'B',None,'B'),(3,'C','B','D'),(4,'D','C','C'),(6,'P','AU','J'),(7,None,None,None),(8,None,None,None),(9,'M',None,None)]:
  pf=fld(P,pcol,N) if pcol else fld(T,'F' if j==7 else 'G',N);mf=fld(M,mcol,E) if mcol else '""';hf=fld(case_tab,hcol,hist_end) if hcol else '""';rf=f'"Check rule: "&INDEX(\'Control Center\'!$A$7:$A$18,{idx})' if j==6 else '""'
  audit.cell(r,j).value=f'=IF({idx}=0,"",IF({kind}="Time",{pf},IF({kind}="Employee",{mf},IF({kind}="History",{hf},{rf}))))'
 for j,review_col in [(5,False),(10,True)]:
  opts=[('Time','Audit' if review_col else T,'V' if review_col else 'A','Review →' if review_col else 'Time row '),('Employee',M,'AU' if review_col else 'B','Fix master' if review_col else 'Employee row '),('History',case_tab,'J' if review_col else 'B','Fix event' if review_col else 'History row '),('Rules','Control Center','B','Fix rule' if review_col else 'Rule row ')]
  expr='""'
  for key,tab,co,lab in reversed(opts):
   row=f'({idx}+{6 if key=="Rules" else 5})';caption=f'"{lab}"' if review_col else f'"{lab}"&{row}'
   expr=f'IF({kind}="{key}",HYPERLINK("#\'{tab}\'!{co}"&{row},{caption}),{expr})'
  audit.cell(r,j).value=f'=IF({idx}=0,"",{expr})'
 # Old scope-index is still used by the fixed rework view; keep source index
 # consistent with the new paged queue for any unhidden inspection.
 audit[f'P{r}']=f'=IF({kind}="Time",{idx},"")'
# Six primary named Excel tables. Embedded reports/rate matrices retain their
# independent auxiliary table filters; they are not additional employee masters.
core_db={'EmployeeMaster':(M,'A5:AU355','One current employee; IQAMA is the lookup key. Edit source fields only.'),'TimesheetTransactions':(T,'A5:BS425','One submitted IQAMA per month-start date. Empty until legitimate timesheets are entered.'),'PayrollResults':(P,'A5:CA425','Calculated settlement, one row per time row. Held net is not payable.'),'MovementHistory':(case_tab,'A5:J425','One event per unique Event ID, linked by IQAMA. Record dated status/site/case changes; evidence is required.'),'RulesConfig':('Control Center','A6:D18','One operating setting per row; includes rule validation. Contract-specific rates remain in the client matrix.'),'AuditExceptions':('Audit','A5:J425','One primary issue per source record on the selected page. Payroll, master, history and config are checked.')}
for table_name,(sh,ref,desc) in core_db.items():
 ws=wb[sh];rg=CellRange(ref);headers=[ws.cell(rg.min_row,j).value for j in range(rg.min_col,rg.max_col+1)];assert all(isinstance(x,str) and x for x in headers) and len(set(x.casefold() for x in headers))==len(headers),(table_name,headers)
 if sh in (M,T,P,'Audit'):ws.auto_filter.ref=None
 tbl=Table(displayName=table_name,ref=ref);tbl.tableStyleInfo=TableStyleInfo(name='TableStyleMedium4',showFirstColumn=False,showLastColumn=False,showRowStripes=False,showColumnStripes=False);ws.add_table(tbl)
# Register capacities are intentionally bounded; Excel's table filter buttons
# replace the overlapping worksheet autofilters and remain enabled on locked views.
# Keep the hidden calculation helper as the only non-user-facing sheet.
jump(hist,'A433',M,'AM5','Edit case fields');jump(g,'A112',case_tab,'A5','Status History  →');g['B112']='Dated status and movement journal; the linked runaway-case view remains at row 431. Source joining dates are imported without reconstructing later changes.'
band(g,157,'Database map · six core tables')
for r,(name,(sheet,ref,desc)) in enumerate(core_db.items(),158):
 jump(g,f'A{r}',sheet,ref.split(':')[0],name);merge(g,r,2,8,desc,INK,PALE if r%2 else WHITE,11);g.row_dimensions[r].height=26
for r,lab,text in [(165,'History workflow','Record a documented change in Status History with a unique reference, date, employee IQAMA, from/to values and evidence. Then update the current master. Editing the master does not create an automatic historical snapshot.'),(166,'Imported history',f'{len(join_events)} dated joins were available in the attached employee records. No site at joining, transfer date, departure date or case date has been inferred. Add events only from supporting records.'),(167,'Exception scope','Payroll: submitted month and dashboard filters. Master: current roster filters. History: effective month and client at event. Rules: always included. All periods ignores all filters for review. Fix the first issue, recalculate and recheck.'),(168,'Review before removing','The six primary tables and hidden Slicer Data feed live formulas. Rate cards and client statement helpers are required for pay; claims and the yearbook are linked. Samples and received evidence are outside live totals, not redundant employee records.'),(169,'Confidential package','The ZIP contains this workbook and a small review guide only. It is not encrypted. Restrict access because the workbook contains employee identifiers, bank details, attendance and pay.')]:guide(r,lab,text)
g['B53']='The dashboard and six data workspaces share one employee master. Status History logs documented events; its case view shows only explicitly reported cases. Slicer Data is a hidden calculation helper.';g['B75']='MovementHistory stores dated event entries; it is not an automatic or immutable change log. The audit queue is a live list of current findings. Signed time corrections and the original 73 submitted rows are retained separately.'
jump(g,'E3','Control Center','A157','Database map');jump(d,'N23','Audit','A5','All exceptions →')
master['A1']='Employee master';master['A2']='One current record per IQAMA. Use the table filters for client and status; record dated changes in Status History. Contract and claim sections stay separate below.'
# Lead with headcount, overdue documents, upcoming renewals and time holds.
# Keep the lower card distinct instead of counting total active twice.
d['B14']='Active employees';d['B15']='='+counte('J','Active');d['B16']='Current workforce in scope';d['F14']='Expired Iqamas';d['F15']='=J108';d['F16']='Documents needing renewal';d['J14']='Due within 15 days';d['J15']="='Slicer Data'!T3";d['J16']='Includes documents due today';d['N14']='Time entries on hold';d['N16']='Resolve before release'
d['B28']='Audit findings';d['B29']="=SUM('Slicer Data'!AB2:AB5)";d['B108']='=IFERROR(B15/J20,"—")';d['J62']=d['J62'].value.replace('B29','B15');d['B17']='=IF(J20=0,"No employees match these filters.","Active staff: "&B15&"   ·   Overdue Iqamas: "&F15&"   ·   Due in 0–15 days: "&J15&"   ·   Missing dates: "&N29)'
for col,fg,bg in [('B',OK_TEXT,MINT),('F',BAD_TEXT,ROSE),('J',WARN_TEXT,YELLOW),('N',BAD_TEXT,ROSE)]:
 for r in (14,15,16):style(d[f'{col}{r}'],fg,bg,23 if r==15 else 11,False)
d['B29'].comment=Comment('Open findings in the current Audit review mode, including master, time, movement and rule issues. Counts primary findings, not people.', 'Basma / Operations')
if 'B29' in [str(x.sqref) for x in d.conditional_formatting]:del d.conditional_formatting['B29']
ink_rule(d,'B29','AND(ISNUMBER(B29),B29>0)',BAD_TEXT)
# Include a repair instruction for the additional configuration hold.
g['A98']='Check payroll settings';merge(g,98,2,8,'Correct the flagged payroll policy in RulesConfig before releasing this line.',INK,PALE,11);nm('RepairRules',"'Control Center'!$A$78:$B$98")
# A single click takes priority work into a correctly scoped consolidated queue.
d['N23'].comment=Comment('Includes payroll holds, master-data exceptions, movement findings and config errors. It is broader than the time-hold KPI.','Basma / Operations')

# Check all dependency surfaces before recalculation, not only formula caches.
for sh in wb:
 for c in sh._cells.values():
  txt=c.value if c.data_type=='f' and isinstance(c.value,str) else ''
  if c.hyperlink and c.hyperlink.location:txt+=' '+c.hyperlink.location
  assert not any("'"+k+"'!" in txt for k in removed),(sh.title,c.coordinate,txt)
for n in wb.defined_names.values():assert not any("'"+k+"'!" in n.attr_text for k in removed)

# Build validation and migration checks performed before calculation.
assert len(ep)==203 and len(cards)==16
ids={cleanid(reg.cell(r,4).value):r for r in range(6,79)};assert len(ids)==73
olds=0
for r in range(4,56):
 if inp.cell(r,3).value:
  olds+=1;rr=ids[cleanid(inp.cell(r,4).value)];assert all(inp.cell(r,8+j).value==reg.cell(rr,6+j).value for j in range(31))
assert olds==50
for s in wb:
 for c in s._cells.values():
  if c.data_type=='f' and isinstance(c.value,str):
   assert len(c.value)<8192 and '#REF!' not in c.value and 'DUMMYFUNCTION' not in c.value,(s.title,c.coordinate)
# Recalculate independently in a temporary profile (no server). Cache transfer avoids office-suite chart/style conversion.
NS={'x':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};q=lambda tag:'{'+NS['x']+'}'+tag
out=ROOT/'Basma_Workforce_and_Payroll.xlsx'
from decimal import Decimal,ROUND_HALF_UP
round2=lambda v:float(Decimal(str(v)).quantize(Decimal('0.01'),rounding=ROUND_HALF_UP))
def recalc(f,dir):
 dir.mkdir(exist_ok=True,parents=True);res=subprocess.run(['soffice',f'-env:UserInstallation={(dir/"profile").as_uri()}','--headless','--convert-to','xlsx','--outdir',str(dir),str(f)],capture_output=True,text=True,timeout=160);print(res.stdout.strip(),res.stderr.strip(),flush=True);target=dir/f.name
 if res.returncode!=0 or not target.is_file():raise RuntimeError('Recalculation failed')
 return target

def attach_cache(original,calculated,outfile):
 with zipfile.ZipFile(original) as base,zipfile.ZipFile(calculated) as calc,zipfile.ZipFile(outfile,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as outz:
  ss=[]
  if 'xl/sharedStrings.xml' in calc.namelist():ss=[''.join(si.itertext()) for si in ET.fromstring(calc.read('xl/sharedStrings.xml'))]
  for f in base.namelist():
   data=base.read(f)
   if re.fullmatch(r'xl/worksheets/sheet\d+.xml',f):
    b=ET.fromstring(data);cr=ET.fromstring(calc.read(f));cells={v.get('r'):v for v in cr.xpath('//x:c',namespaces=NS)}
    for cell in b.xpath('//x:c[x:f]',namespaces=NS):
     dest=cell.find(q('v'));cc=cells.get(cell.get('r'))
     if cc is None:continue
     val=cc.find(q('v'));typ=cc.get('t','n');vt=val.text if val is not None else None
     if typ=='s':vt=ss[int(vt)];typ='str'
     if typ=='inlineStr':vt=''.join(cc.find(q('is')).itertext());typ='str'
     if typ=='e':raise AssertionError((f,cell.get('r'),vt,cell.find(q('f')).text))
     if dest is None:dest=ET.SubElement(cell,q('v'))
     cell.set('t',typ);dest.text=vt
    data=ET.tostring(b,encoding='UTF-8')
   outz.writestr(f,data)

def validate(path):
 vv=xl.load_workbook(path,data_only=True); p=vv[P];t=vv[T];dash=vv['Dashboard'];counts=collections.Counter(p[f'O{r}'].value for r in range(6,79));print('Release rows:',counts,'Issues:',collections.Counter(p[f'P{r}'].value for r in range(6,79)),flush=True)
 assert vv.sheetnames==order;assert dash['B20'].value==0 and dash['J20'].value==203
 # Independently validate every new KPI from the sole employee-input sheet and ranked ledger.
 md=vv['Main Dataset'];asof=vv['Control Center']['B10'].value.date()
 remaining=[]
 for r in range(6,E+1):
  if md[f'B{r}'].value and isinstance(md[f'D{r}'].value,dt.datetime):remaining.append((md[f'D{r}'].value.date()-asof).days)
 for row,i,lo,hi,_ in windows:assert dash.cell(row+1,starts[i]).value==sum(lo<=days<=hi for days in remaining),(row,i,lo,hi)
 assert dash['B32'].value==sum(v[9] in ['Suspended',*departed_states] for v in ep)
 assert dash['F32'].value==sum(isinstance(p[f'Z{r}'].value,(int,float)) for r in range(6,N+1))==0
 assert dash['J20'].value==dash['B15'].value+dash['F108'].value+dash['B32'].value
 # Check the added analytical chart data independently of the dashboard formulas.
 helper=vv['Slicer Data'];age_check=[sum(x<0 for x in remaining)]+[sum(a<=x<=b for x in remaining) for a,b in [(0,15),(16,30),(31,60)]]+[sum(x>60 for x in remaining),len(ep)-len(remaining)]
 assert [helper[f'T{r}'].value for r in range(2,8)]==age_check
 stat_check=collections.Counter(v[9] for v in ep);status_counts=[stat_check[st] for st in ['Active','On Leave','Suspended']]+[sum(stat_check[st] for st in departed_states)];status_counts.append(len(ep)-sum(status_counts))
 assert [helper[f'W{r}'].value for r in range(2,7)]==status_counts
 assert sum(age_check)==sum(status_counts)==dash['J20'].value
 assert sum(helper[f'Z{r}'].value for r in range(2,6))==dash['F20'].value
 assert vv['Control Center']['B67'].value==30 and vv['Control Center']['B66'].value==dt.datetime(2026,9,30)
 for r in range(6,N+1):
  for k,j in enumerate(trace_cols,38):assert p.cell(r,k).value==t.cell(r,j).value,(r,k,j)
  assert md.cell(r+357,7).value==p[f'O{r}'].value and md.cell(r+357,9).value==vv['Audit'][f'S{r}'].value
 assert len(join_events)==70 and len({vv[case_tab][f'E{r}'].value for r in range(6,76)})==70
 assert all(vv[case_tab][f'J{r}'].value=='Recorded' for r in range(6,76))
 assert all(vv[case_tab][f'J{r}'].value is None for r in range(76,hist_end+1))
 assert all(vv[case_tab].cell(i+6,2).value==event[2] and vv[case_tab].cell(i+6,3).value==event[1] for i,event in enumerate(join_events))
 assert dash['B15'].value==sum(v[9]=='Active' for v in ep) and dash['F15'].value==age_check[0] and dash['J15'].value==age_check[1]
 assert dash['B29'].value==17 and helper['AB3'].value==17 and helper['AB2'].value==helper['AB4'].value==helper['AB5'].value==0
 assert all(vv['Audit'][f'F{r}'].value=='Check master salary' for r in range(6,23)) and vv['Audit']['F23'].value is None
 # Verify an empty live intake independently of the retained imported evidence.
 assert all(t.cell(r,j).value is None for r in range(6,N+1) for j in [*rawcols,67])
 assert all(t[f'BN{r}'].value==0 and t[f'E{r}'].value is None and p[f'O{r}'].value is None for r in range(6,N+1))
 assert all(dash[ad].value==0 for ad in ['S6','S7','B20','F20','F32']) and p['J433'].value==0 and t['F432'].value==0
 archive=vv['Audit'];assert sum(archive.cell(r,1).value!='Checks passed' for r in range(433,506))==18
 for r in range(6,79):
  ar=r+427
  assert archive.cell(ar,3).value==reg.cell(r,1).value and archive.cell(ar,4).value==reg.cell(r,3).value and str(archive.cell(ar,5).value)==cleanid(reg.cell(r,4).value)
  for day in range(30):assert archive.cell(ar,6+day).value==reg.cell(r,6+day).value
 assert collections.Counter(md[f'A{r}'].value for r in range(6,209))==collections.Counter(v[6] for v in ep)
 assert [md[f'A{r}'].value for r in range(1103,1143)]==[md[f'A{r}'].value for r in range(793,833)]
 assert sum(md[f'K{r}'].value or 0 for r in range(1103,1143))==203
 # Check the two explicit legacy reports, without inventing clients or months.
 case_rows=[r for r in range(case_start,case_end+1) if vv[case_tab][f'E{r}'].value];source_cases=[r for r in range(6,E+1) if md[f'AM{r}'].value=='Reported']
 assert len(case_rows)==len(source_cases)==2 and vv[case_tab]['E433'].value==2
 for i,j in zip(case_rows,source_cases):
  assert vv[case_tab][f'C{i}'].value==md[f'B{j}'].value and vv[case_tab][f'D{i}'].value==md[f'C{j}'].value
  assert vv[case_tab][f'A{i}'].value=='Not recorded' and vv[case_tab][f'B{i}'].value=='Not recorded' and vv[case_tab][f'E{i}'].value=='Reported'
 assert dash['B111'].value==sum(v[9] in departed_states for v in ep) and dash['F111'].value==dash['B15'].value
 assert dash['J111'].value==sum(md[f'V{r}'].value==md[f'W{r}'].value=='Current' for r in range(6,E+1)) and dash['N111'].value==age_check[3]
 assert vv['Audit']['D518'].value=='Terminated' and vv['Audit']['B519'].value is None
 assert not any(c.data_type=='e' for ws in vv for c in ws._cells.values());print('Verified empty live intake and payroll, retained all 73 archived submissions, all 203 client-first employee records and 10 expiry windows.',flush=True)
 return vv

with tempfile.TemporaryDirectory(prefix='basma-build-') as temp:
 td=Path(temp);fq=td/'model.xlsx';wb.save(fq);calc=recalc(fq,td/'calculated');cv=xl.load_workbook(calc,data_only=True);print('Dashboard preview:',[(x,cv['Dashboard'][x].value) for x in ['B15','F15','J15','N15','B20','F20','J20','N20']],flush=True)
 errors=[(sh.title,c.coordinate,c.value) for sh in cv for c in sh._cells.values() if c.data_type=='e'];print('Calculation errors:',len(errors),errors[:10],flush=True)
 if errors:raise AssertionError(errors[:10])
 # Seed native chart caches to make previews and PDF readers show the same values before recalculation.
 for chart,catcol,valcols,end in [(bc,'K','LM',7),(dc,'P','Q',3),(aging_chart,'S','T',7),(status_chart,'V','W',6),(flowchart,'Y','Z',5)]:
  sh=cv['Slicer Data'];cats=[sh[f'{catcol}{r}'].value for r in range(2,end+1)]
  for ser,vcol in zip(chart.series,valcols):
   ser.cat=AxDataSource(strRef=StrRef(f=f"'Slicer Data'!${catcol}$2:${catcol}${end}",strCache=StrData(ptCount=len(cats),pt=[StrVal(idx=i,v=str(v or '')) for i,v in enumerate(cats)])))
   vals=[sh[f'{vcol}{r}'].value for r in range(2,end+1)];ser.val.numRef.numCache=NumData(formatCode='0.00',ptCount=len(vals),pt=[NumVal(idx=i,v=val) for i,val in enumerate(vals) if isinstance(val,(int,float))])
 fit_registers(cv);apply_readable_ui();monthly_finish();wb.save(fq);attach_cache(fq,calc,out);validate(out)
# Structural verification of the final Excel package, not just the in-memory model.
check=xl.load_workbook(out);assert check.sheetnames==order and check.active.title=='Dashboard';assert len(check['Dashboard']._charts)==2 and len(check[P]._charts)==3 and len(check['Dashboard'].data_validations.dataValidation)==8
expected_checks={
 case_tab:('A5','Client at event',hist_end),
 'Timesheet_Input':('D5','IQAMA',N), 'Salarysheet':('O5','Release',N), 
 'Main Dataset':('B5','Employee Name',E), 'Audit':('F5','First review item',N),
 'Control Center':('A1','Control center',60),
 'Dashboard':('B3','Workforce action desk',123),'Slicer Data':('A1','Sites',353),
 }
for sh,(addr,expected,last) in expected_checks.items():
 assert check[sh][addr].value==expected and check[sh].max_row>=last,(sh,addr)
assert sum(check['Main Dataset'].cell(r,2).value is not None for r in range(6,E+1))==203
assert sum(check[T].cell(r,3).value is not None for r in range(6,N+1))==0
assert sum(check[M].cell(r+787,2).value is not None for r in range(6,C+1))==16
assert check[T].tables['TimesheetTransactions'].autoFilter.ref==f'A5:BS{N}' and not check[T].protection.autoFilter and check[T]['E3'].hyperlink.location=="'Timesheet_Input'!A430"
assert check[T]['B432'].value=='All' and check[T]['D432'].value=='All' and check[T]['BU854'].data_type=='f'
assert check[T]['A1'].font.b is False and check[M]['A1'].font.b is False and check[M]['A1102'].value=='Client / worksite'
assert [check[case_tab].cell(435,j).value for j in range(1,6)]==['Client','Month / year','Employee','IQAMA','Case status']
assert check[case_tab]['A1'].font.b is False and check[case_tab]['C785'].data_type=='f'
assert check[P]['A5'].value==check[T]['A5'].value=='Client / worksite' and check[P]['B5'].value=='Month / year'
assert check['Audit']['A5'].value=='Client' and check[case_tab]['G435'].value=='Case remark' and check[case_tab]['G785'].data_type=='f'
assert check[T]['A434'].value=='Client / worksite' and check[T]['B434'].value=='Month / year' and check[T]['H434'].value=='Remarks'
assert check[T]['I432'].value=='All periods' and check[T]['I432'].protection.locked==False
assert check[T]['E5'].value=='Identity check' and check[P]['N5'].value=='Ready net' and check[P]['N425'].data_type=='f' and check['Main Dataset']['A5'].value=='Client / worksite';assert 'Salarysheet' in check['Dashboard']['S6'].value
assert check[M]['AF5'].value=='Period entries' and check[M]['AG5'].value=='Submission' and check[M]['V5'].value=='Iqama status'
assert check[P]['C435'].value=='Employee name' and check[P]['J433'].data_type=='f' and check[P]['C855'].data_type=='f' and check[P]['D433'].protection.locked==False
assert check['Control Center']['B123'].value==examples['B1'].value and check['Control Center']['F124'].value=='=TEXT(F125,"ddd")'
assert all(check[M].cell(r,j).data_type=='f' for r in range(6,E+1) for j in range(32,39))
for tab,(top,entry,last,caption) in sample_specs.items():
 assert check[tab][entry].hyperlink.location==f"'{tab}'!A{top}"
 assert check[tab].cell(top,1).value.startswith('Sample row  /  ') and check[tab].cell(top+1,1).value==check[tab].cell(5,1).value
 assert check[tab]['A1'].font.b is False
print(f'Built {out.name} • {out.stat().st_size:,} bytes • 48 KPIs • 5 live charts • verified full time-to-pay transfer and source hub.',flush=True)

# Optional release tests. All edits use temporary copies: no generated employee,
# test status, attendance mark or signature is left in the delivered workbook.
if os.environ.get('BASMA_VERIFY')=='1':
 with tempfile.TemporaryDirectory(prefix='basma-regression-') as tmp:
  tmp=Path(tmp);test=xl.load_workbook(out);delivered=xl.load_workbook(out,data_only=True)
  def calc_test(stage):
   sub=tmp/stage;sub.mkdir();test.save(sub/'trial.xlsx');v=xl.load_workbook(recalc(sub/'trial.xlsx',sub/'out'),data_only=True)
   assert not any(c.data_type=='e' for sh in v for c in sh._cells.values()),stage
   return v
  roster=test[M];entry=test[T];review=test['Audit'];controls=test[d.title];refdate=dt.datetime(2026,11,19)
  # A separate, deterministic cohort exercises every edge and both missing
  # report states, without attributing any actual departure date to a worker.
  delta=[-1,0,1,2,8,9,10,13,14,15,16,19,20,21,29,30,31,59,60,61,None,0]
  cohort_status=['Active','On Leave','Suspended','Terminated','Left','Resigned']
  for i,n in enumerate(delta):
   r=6+i;roster[f'H{r}']='TEST COHORT';roster[f'J{r}']=cohort_status[i%6];roster[f'D{r}']=None if n is None else 0 if i==21 else refdate+dt.timedelta(days=n)
  roster['R6']=0;test['Control Center']['B10']=refdate;controls['J9']='TEST COHORT';test[M]['A832']='TEST END SITE'
  bounds={'J32':(9,15),'N32':(1,14),'B38':(0,60),'F38':(0,30),'J38':(0,20),'N38':(0,15),'B42':(0,14),'F42':(0,9),'J42':(0,1),'N42':(0,0)}
  v=calc_test('roster');dv0=v['Dashboard'];assert dv0['J20'].value==len(delta)
  for addr,(lo,hi) in bounds.items():assert dv0[addr].value==sum(lo<=n<=hi for n in delta[:-1] if isinstance(n,int)),(addr,lo,hi,dv0[addr].value)
  counts=collections.Counter(cohort_status[i%6] for i in range(len(delta)));gone=sum(counts[x] for x in departed_states)
  assert dv0['B111'].value==gone and dv0['B15'].value==counts['Active'] and dv0['B32'].value==gone+counts['Suspended']
  assert v['Slicer Data']['W5'].value==gone and dv0['N29'].value==2 and v[M]['W6'].value=='Missing date'
  assert all(v['Audit'][f'D{r}'].value in departed_states for r in range(518,518+gone));assert v['Audit'][f'A{518+gone}'].value is None
  assert v[case_tab]['E433'].value==2 and any(v['Slicer Data'][f'A{r}'].value=='TEST END SITE' for r in range(3,43))
  expected_days=sorted(n for n in delta[:-1] if isinstance(n,int) and 0<=n<=60)[:10]
  assert [v['Dashboard'][f'L{rr}'].value for rr in range(67,67+len(expected_days))]==expected_days
  controls['M65']='Already expired';v=calc_test('overdue');assert v['Dashboard']['L67'].value==-1 and v['Dashboard']['N67'].value=='Renew overdue' and v['Dashboard']['B68'].value is None
  controls['M65']='Missing date';v=calc_test('undated');assert v['Dashboard']['N67'].value==v['Dashboard']['N68'].value=='Record expiry' and v['Dashboard']['L67'].value is None and v['Dashboard']['B69'].value is None
  print('PASS: all three renewal queues follow the active roster scope and sort by actual expiry date; missing dates are not overdue.')
  print('PASS: source-linked status list, all 10 expiry boundaries, missing/zero dates, departed vs inactive vs runaway, new client across a blank gap.')
  # Start the payroll test from the unmodified delivered file. Copy actual
  # archived submissions back with client first and month on each input row.
  test=xl.load_workbook(out);entry=test[T];review=test['Audit'];controls=test['Dashboard'];archive=delivered['Audit'];srows=[(6,433),(7,434),(8,435),(9,436),(10,437),(11,438),(12,438),(13,439),(14,440),(15,441)]
  for dest,received_row in srows:
   for j,k in [(1,3),(2,2),(3,4),(4,5),*[(8+i,6+i) for i in range(30)]]:entry.cell(dest,j).value=archive.cell(received_row,k).value
  entry['C7']='UNVERIFIED INPUT';entry['A8']='PLASTIC';entry['H9']='X';entry['AK10']=None;entry['B13']=dt.datetime(2026,9,2);entry['D14']='12*';entry['AL15']='P';entry['B432']='SHETA & SAIF';entry['D432']='Hold';review['D3']='All periods'
  # The folded claim register remains the sole writable expense source.
  test[M]['A1092']=dt.datetime(2026,9,8);test[M]['D1092']='Travel';test[M]['E1092']=123.45;test[M]['F1092']='Approved';test[M]['G1092']='Cash'
  # Record an actual before-correction review so resolution must be signed.
  review['V7']='TEST REVIEWER';review['W7']=dt.datetime.today().replace(hour=0,minute=0,second=0,microsecond=0);review['X7']='Test fixture: review corrected name against the received document.'
  v=calc_test('paste');expected={6:'Checks passed',7:'Name mismatch',8:'Site mismatch',9:'Invalid day code',10:'Incomplete days',11:'Duplicate entry',12:'Duplicate entry',13:'Invalid period',14:'Invalid IQAMA',15:'Outside month'}
  for r,reason in expected.items():
   assert v[P][f'P{r}'].value==reason,(r,reason,v[P][f'P{r}'].value)
   if reason!='Checks passed':assert v[P][f'O{r}'].value=='Hold' and v[P][f'N{r}'].value==0
  assert abs(v[M]['D839'].value-123.45)<.005 and v[M]['F839'].value==1 and abs(v[M]['K1213'].value-123.45)<.005
  assert v[P]['A6'].value==v[T]['A6'].value and v[P]['B6'].value==v[T]['B6'].value
  assert v[T]['A435'].value=='SHETA & SAIF' and v[T]['B435'].value==dt.datetime(2026,9,1)
  assert v[T]['F435'].value=='Hold' and v['Audit']['A6'].value==v[T]['A7'].value and v['Audit']['E6'].value=='Time row 7'
  assert v[T]['F432'].value==8 and v[P]['N6'].value==1200 and v['Dashboard']['B20'].value==9 and v['Dashboard']['S6'].value==1200
  assert all(v[P].cell(6,38+i).value==v[T].cell(6,8+i).value for i in range(31));assert v['Dashboard']['F32'].value==9
  assert v[P]['AL15'].value==v[T]['H15'].value and any(v['Audit'][f'F{r}'].value=='Invalid period' for r in range(6,20))
  entry['C7']=archive['D434'].value;test[P]['D433']='Hold';v=calc_test('corrected')
  assert v[P]['AJ7'].value=='Checks passed' and v[P]['P7'].value=='Document correction' and v[P]['N7'].value==0 and v[P]['J433'].value==8
  review['Z7']=v['Audit']['Y7'].value;entry['B432']='All';entry['D432']='Ready';v=calc_test('signed')
  assert v['Audit']['S7'].value=='Remade' and v[P]['N7'].value==1200 and v['Dashboard']['S6'].value==2400 and v[T]['F432'].value==2
  entry['H7']='A';entry['B6']=dt.datetime(2028,2,1);entry['AK6']=None;controls['B9']=dt.datetime(2028,2,1);controls['J12']='Absence days';v=calc_test('changed')
  assert v[P]['P7'].value=='Document correction' and v[P]['N7'].value==0 and v['Audit']['AD7'].value==0
  assert v[T]['BF6'].value==29 and v[P]['O6'].value=='Ready' and v['Dashboard']['B20'].value==1 and v['Dashboard']['B124'].value==1 and v['Control Center']['B67'].value==29
  # Test appended events without mutating the shipped source or inferring changes.
  test=xl.load_workbook(out);history=test[case_tab];review=test['Audit'];review['D3']='All periods';master_test=test[M];empid=master_test['C6'].value
  event_day=dt.datetime(2026,9,1)
  seeds=[('QA-001','Client transfer','Site A','Site B','Signed transfer notice'),('QA-002','Status change','Active','On Leave','Approved leave record'),('QA-002','Joining',None,None,'Test duplicate ID'),('QA-004','Exit',None,'Active','Test rejected exit state'),('QA-005','Correction','Earlier value','Corrected value',None),('QA-006','Client transfer','Site B','Site C','Test missing master'),('QA-007','Case update',None,'Reported','Test date'),('QA-008','Joining',None,None,'Source row A'),('QA-009','Joining',None,None,'Source row B')]
  for rr,(key,typ,prior,nxt,evidence) in enumerate(seeds,76):
   for col,val in {1:'Site A',2:event_day,3:empid,5:key,6:typ,7:prior,8:nxt,9:evidence}.items():history.cell(rr,col).value=val
  history['C81']='9999999999';history['B82']=None
  test['Control Center']['B11']='invalid';test['Control Center']['B13']='invalid'
  for j,k in [(1,3),(2,2),(3,4),(4,5),*[(8+i,6+i) for i in range(30)]]:test[T].cell(6,j).value=delivered['Audit'].cell(433,k).value
  v=calc_test('history-controls');findings=[v[case_tab][f'J{rr}'].value for rr in range(76,85)]
  assert findings==['Recorded','Duplicate event ID','Duplicate event ID','Check exit status','Evidence missing','Check employee ID','Check effective date','Duplicate event','Duplicate event'],findings
  assert v[P]['P6'].value=='Check payroll settings' and v[P]['N6'].value==0
  assert v['Slicer Data']['AB4'].value==8 and v['Slicer Data']['AB5'].value==2 and v['Dashboard']['B29'].value==28
  issues={v['Audit'][f'F{rr}'].value for rr in range(6,50)};assert {'Duplicate event','Duplicate event ID','Evidence missing','Check payroll settings','Check rule: Weekly premium cycle','Check rule: Half-day absence charge'}<=issues
  test['Control Center']['B11']=7;test['Control Center']['B13']=0
  # More than a page of findings: the last page must remain addressable.
  for rr in range(76,426):history[f'E{rr}']='QA-DUP'
  for rr in range(6,106):test[M][f'P{rr}']=None
  review['G3']=2;v=calc_test('exception-pagination');total=sum(v['Slicer Data'][f'AB{j}'].value for j in range(2,6));assert total>420
  assert v['Audit']['F6'].value and sum(bool(v['Audit'][f'F{rr}'].value) for rr in range(6,426))==total-420
  assert v[case_tab]['E433'].value==2 and v[case_tab]['B6'].value==join_events[0][2]
  print('PASS: six linked HR tables, 70 preserved historical joins, duplicate/missing/orphan movement checks, invalid-configuration payroll hold, and a multi-page exception queue.')
  print('PASS: paste-to-payroll, eight blocking conditions, scoped and all-period audits, filters, ranking, leap February, signed correction and invalidated signature after an edit.')
 assert len(list(ROOT.glob('*.xlsx')))==1 and sum(sh.sheet_state=='visible' for sh in check)==7
 assert all(check[M].cell(rr,1).protection.locked is False for rr in (793,832,843,1092))
 assert [k for k in check[M].tables]==['ContractMatrix','ClaimRegister','WorksiteSummary','EmployeeMaster']
 assert check['Dashboard']['B67'].data_type=='f' and check['Dashboard']['M65'].value=='Next 60 days' and not check[T]['A6'].value
 with zipfile.ZipFile(out) as z:assert z.testzip() is None
 assert all(not c.value or c.data_type=='f' for c in check[T]['A'][5:N]);assert len(check['Dashboard']._charts)==2 and len(check[P]._charts)==3 and len(check.sheetnames)==8
 for sh in check:
  for c in sh._cells.values():
   if c.hyperlink and c.hyperlink.location:
    tab,dest=c.hyperlink.location.rsplit('!',1);assert tab.strip("'") in check.sheetnames,(sh.title,c.coordinate,tab)
 print('PASS: final workbook is clean; all tabs, chart sources, case records, entry controls and internal navigation retained.')

# Final targeted layout pass. Keep the six core tables, but co-locate four
# formerly independent workspaces. No parallel data sources or hidden tabs.
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.pagebreak import Break
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.chart import BarChart
from functools import lru_cache
op_name='Operations Control';op_offsets={'Audit':64,'Status History':950,'Control Center':1760,'Slicer Data':2000}
old_tabs=tuple(op_offsets);tag_re='|'.join(map(re.escape,old_tabs))
ref_move=re.compile(r"'("+tag_re+r")'!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)\b")
def move_xref(f):return ref_move.sub(lambda m:f"'{op_name}'!"+relocated(m[2],{},op_offsets[m[1]]),f)
def op_addr(old,addr):return relocated(addr,{},op_offsets[old])
def move_link(loc):
 for tab in old_tabs:
  if loc.startswith("'"+tab+"'!"):return f"'{op_name}'!"+op_addr(tab,loc.split('!',1)[1])
 return loc
# Dynamic HYPERLINK targets are text, not normal formula reference tokens.
# Rewrite each supported address shape explicitly; never change the display
# text or a time-row reference because a reviewing sheet has moved.
linked_index_pattern=re.compile(r'(HYPERLINK\("#\'('+tag_re+r")'!([A-Z]+)\"&\()('Slicer Data'!\$?[A-Z]+\$?\d+)(\+)(\d+)(\),)")
staticparts=re.compile(r'''"#'('''+tag_re+r''')'!"&"([A-Z]+)(\d+)"''')
@lru_cache(maxsize=8192)
def op_formula(f,source=''):
 if not isinstance(f,str):return f
 dy=op_offsets.get(source,0)
 if not dy and not any("'"+k+"'!" in f for k in old_tabs):return f
 # Literal/chained hyperlinks from the old client workspace have been replaced
 # below; the other dynamic links use an explicit matched-row offset.
 f=linked_index_pattern.sub(lambda m:'HYPERLINK("#\''+op_name+"'!"+m[3]+'"&('+m[4]+'+'+str(int(m[6])+op_offsets[m[2]])+m[7],f)
 f=staticparts.sub(lambda m:'"#\''+op_name+"'!"+m[2]+str(int(m[3])+op_offsets[m[1]])+'"',f)
 if dy:
  prefix=f.startswith('=');f=local_formula(f if prefix else '='+f,{},dy);f=f if prefix else f[1:]
  # Rebased ROW() is needed for MATCH/sequence formulas in packed views.
  f=re.sub(r'\bROW\(\)-(\d+)',lambda m:'ROW()-'+str(int(m[1])+dy),f)
 return move_xref(f)

# Use the fully validated existing workbook, including final fonts/widths.
base_view=xl.load_workbook(out,data_only=True);stage=xl.load_workbook(out);hub=stage.create_sheet(op_name);setup(hub);hub.sheet_view.zoomScale=90;hub.sheet_format.defaultRowHeight=18;hub.sheet_properties.tabColor=TEAL;hub.sheet_properties.outlinePr.summaryBelow=False
# Each section has its own table header / filter. Shared columns use widths
# suitable for names, dates and notes; do not hide a column required elsewhere.
for j in range(1,79):hub.column_dimensions[L(j)].width=15 if j>11 else [26,18,31,21,23,19,24,29,27,19,24][j-1]
for col,wd in [('R',18),('S',20),('T',24),('U',43),('V',21),('W',18),('X',38),('Y',29),('Z',29)]:hub.column_dimensions[col].width=wd
# All externally triggered formula jumps to the live audit queue follow its
# new position, including client-level navigation when no record matches.
for rr in range(1103,1143):
 stage[M][f'F{rr}']=f'=IF(A{rr}="","",HYPERLINK("#\'{op_name}\'!"&IFERROR("F"&(MATCH(A{rr},\'{op_name}\'!$A$70:$A$489,0)+69),"D67"),"Review"))'
for tab,dy0 in op_offsets.items():
 srcs=stage[tab]
 for cc in srcs._cells.values():
  if cc.value is None and not getattr(cc,'has_style',False):continue
  if isinstance(cc,xl.cell.cell.MergedCell):continue
  dest=hub.cell(cc.row+dy0,cc.column);dest.value=op_formula(cc.value,tab) if cc.data_type=='f' else cc.value;dest._style=copy.copy(cc._style)
  if cc.comment:dest.comment=copy.copy(cc.comment)
  if cc.hyperlink:
   lk=copy.copy(cc.hyperlink);lk.ref=dest.coordinate
   if lk.location:lk.location=move_link(lk.location)
   dest.hyperlink=lk
 for rg in srcs.merged_cells.ranges:
  ar=CellRange(str(rg));ar.shift(row_shift=dy0);hub.merge_cells(str(ar))
 for row,dim in srcs.row_dimensions.items():
  rd=hub.row_dimensions[row+dy0];rd.height=dim.height;rd.hidden=dim.hidden;rd.outlineLevel=dim.outlineLevel;rd.collapsed=dim.collapsed
 for val in srcs.data_validations.dataValidation:
  v=copy.deepcopy(val);new_ranges=[]
  for rg in val.sqref.ranges:
   ar=CellRange(str(rg));ar.shift(row_shift=dy0);new_ranges.append(str(ar))
  v.sqref=' '.join(new_ranges)
  if v.formula1:v.formula1=op_formula(v.formula1,tab)
  if v.formula2:v.formula2=op_formula(v.formula2,tab)
  hub.add_data_validation(v)
 for cf in srcs.conditional_formatting:
  new_ranges=[]
  for rg in cf.sqref.ranges:
   ar=CellRange(str(rg));ar.shift(row_shift=dy0);new_ranges.append(str(ar))
  for rule in srcs.conditional_formatting[cf]:
   nr=copy.deepcopy(rule);nr.formula=[op_formula(f,tab) for f in rule.formula];hub.conditional_formatting.add(' '.join(new_ranges),nr)
 for table in list(srcs.tables.values()):
  nt=copy.deepcopy(table);ar=CellRange(nt.ref);ar.shift(row_shift=dy0);nt.ref=str(ar)
  if nt.autoFilter:nt.autoFilter.ref=str(ar)
  del srcs.tables[table.name];hub.add_table(nt)
for sh in stage:
 if sh.title in old_tabs or sh==hub:continue
 for cc in sh._cells.values():
  if cc.data_type=='f' and isinstance(cc.value,str):cc.value=op_formula(cc.value)
  if cc.hyperlink and cc.hyperlink.location:cc.hyperlink.location=move_link(cc.hyperlink.location)
 for cf in sh.conditional_formatting:
  for rule in sh.conditional_formatting[cf]:rule.formula=[op_formula(f) for f in rule.formula]
 for val in sh.data_validations.dataValidation:
  if val.formula1:val.formula1=op_formula(val.formula1)
  if val.formula2:val.formula2=op_formula(val.formula2)
 for ch in sh._charts:
  for ser in ch.series:
   for target in [getattr(ser.tx,'strRef',None),getattr(ser.val,'numRef',None),getattr(ser.cat,'strRef',None),getattr(ser.cat,'numRef',None)]:
    if target and target.f:target.f=op_formula(target.f)
for n in stage.defined_names.values():n.attr_text=op_formula(n.attr_text)
for tab in old_tabs:stage.remove(stage[tab])
# A settings-only edit is now on a protected page: unlock only the actual
# input slots and keep guides, report formulas, archives and helpers locked.
view_lock(hub);hub.protection.formatRows=False
for rr in range(7,19):hub[f'B{rr+op_offsets["Control Center"]}'].protection=unlock
for rr in range(436,786):hub[f'F{rr+op_offsets["Status History"]}'].number_format=';;;'
for rr in range(6,426):
 for col in ['P','Q','AC','AD']:hub[f'{col}{rr+op_offsets["Audit"]}'].number_format=';;;'
 hub[f'K{rr+op_offsets["Status History"]}'].number_format=';;;'
# Header naming is aligned by meaning. Historical client at event is not
# falsely treated as the current assignment; the core table keys stay intact.
for ref,text in [('A65','Audit & exceptions'),('A951','Status & movement history'),('A1761','Control center'),('A1765','Operating settings'),('A2001','Sites')]:hub[ref]=text
for a,text in [('B1813','Five tabs: Dashboard, Main Dataset, Timesheet_Input, Salarysheet and Operations Control. This tab contains workload pressure, audit, movement history, settings and internal calculation data.'),('B1864','Time transactions: paste approved source data in the time register, not on this control page.'),('A1860','Workbook map · five visible tabs'),('A1871','Calculation helpers'),('B1871','Grouped rows 2001 onward contain dropdowns, chart data, rankings and exception indices. The formulas remain live when this detail is collapsed.'),('B1872','Status History is the dated movement journal at row 951. The linked reported-case list begins at row 1381; no new event has been inferred.'),('B1925','Append a documented event in the MovementHistory table on this sheet, then update the master. Keep the actual date, unique ID, prior/new value and supporting evidence.'),('B1928','Keep all six core tables. Dropdowns and charts use the collapsed calculation section of this sheet; do not delete it. Samples and archived records never feed live payroll.')]:hub[a]=text
# Bring back-to-home links and a permanent row of section anchors. Long
# evidence/history tables remain filterable without permanently freezing a
# tall band of instructions over a small laptop screen.
merge(hub,1,1,8,'Operations control',WHITE,NAVY,20,False);hub.row_dimensions[1].height=34
merge(hub,2,1,8,'Workload, exceptions, movements and rules in one workspace. Blue cells are inputs; calculated cells are protected.',INK,PALE,12);hub.row_dimensions[2].height=25
for c1,c2,label,dst in [(1,2,'Workload pressure','A5'),(3,4,'Audit & rework','A65'),(5,6,'Status history','A951'),(7,8,'Control center','A1761')]:
 cell=merge(hub,3,c1,c2,label,BLUE,SKY,12,False);cell.hyperlink=Hyperlink(ref=cell.coordinate,location=f"'{op_name}'!{dst}",display=label);cell.border=Border(bottom=Side(style='thin',color='7DA4AA'))
hub.row_dimensions[3].height=26;hub.freeze_panes='A4';hub.sheet_view.selection[-1].activeCell='A5';hub.sheet_view.selection[-1].sqref='A5'
for addr in ['A67','A953','A1763']:jump(hub,addr,op_name,'A1','↑ Operations home')
# Explicit cross-section routes so no user needs to know the long row offsets.
jump(hub,'B953',op_name,'A1376','Sample row');jump(hub,'C953',op_name,'A1381','Reported cases');jump(hub,'E67',op_name,'R68','Rework log');jump(hub,'J67',op_name,'A579','Departures');jump(hub,'E1763',op_name,'A1917','Database map')
for r in (1867,1868,1869):
 ce=hub[f'B{r}'];ce.value=str(ce.value or '').replace('not a separate tab.','not a separate tab.')
# Workload is an operational task measure, not a score on employees. The
# three roster tasks use current site; held timesheets use submitted site.
# Tasks can overlap. Single-site formulas plus an unattributed bucket exactly
# reconcile to global totals, including pasted unknown worksite values.
band(hub,5,'Workload pressure · priority by worksite',1,8);hub.row_dimensions[5].height=25
merge(hub,6,1,8,'="Reporting month: "&TEXT(ViewMonth,"mmm yyyy")&"     /     Worksite: "&ViewSite&"     /     Compliance date: "&TEXT(ReviewDate,"dd mmm yyyy")&"     /     Renewal alert: "&IqamaWarnDays&" days"',INK,PALE,11);hub.row_dimensions[6].height=23
base_helpers=2900;ph=base_helpers+1;pend=base_helpers+41
hdrs=['Client / worksite','Missing timesheets','Expired Iqamas','Upcoming renewals','Held time rows','Open tasks','Immediate tasks','Priority order']
for j,txt in enumerate(hdrs,1):hub.cell(base_helpers,j,txt)
mr=lambda col:rng(M,col,E);pr=lambda col:rng(P,col,N)
for r in range(ph,pend):
 srcrow=r-ph+793;site=f"IF(ISBLANK('{M}'!A{srcrow}),\"\",'{M}'!A{srcrow})";hub[f'A{r}']=f'=IF(OR({site}="",COUNTIF(\'{M}\'!$A$793:A{srcrow},{site})>1,AND(ViewSite<>"All",ViewSite<>{site})),"",{site})'
 # Both ends of the dynamic card-list range must refer to Main Dataset.
 hub[f'A{r}']=hub[f'A{r}'].value.replace(f"'!$A$793:A{srcrow}",f"'!$A$793:$A{srcrow}")
 base=f'{mr("Y")},1,{mr("A")},A{r}';fs=[f'COUNTIFS({base},{mr("J")},"Active",{mr("AF")},0)',f'COUNTIFS({base},{mr("V")},"Expired")',f'COUNTIFS({base},{mr("V")},"Due soon")',f'COUNTIFS({pr("V")},1,{pr("A")},A{r},{pr("O")},"Hold")']
 for j,formula in enumerate(fs,2):hub.cell(r,j,f'=IF(A{r}="",0,{formula})')
r=pend;hub[f'A{r}']='Unlisted / no site'
for col,total in [('B',f'COUNTIFS({mr("Y")},1,{mr("J")},"Active",{mr("AF")},0)'),('C',f'COUNTIFS({mr("Y")},1,{mr("V")},"Expired")'),('D',f'COUNTIFS({mr("Y")},1,{mr("V")},"Due soon")'),('E',f'COUNTIFS({pr("V")},1,{pr("O")},"Hold")')]:hub[f'{col}{r}']=f'={total}-SUM({col}{ph}:{col}{r-1})'
for r in range(ph,pend+1):
 hub[f'F{r}']=f'=SUM(B{r}:E{r})';hub[f'G{r}']=f'=C{r}+E{r}';hub[f'H{r}']=f'=IF(F{r}=0,"",COUNTIF($G${ph}:$G${pend},">"&G{r})+COUNTIFS($G${ph}:$G${pend},G{r},$F${ph}:$F${pend},">"&F{r})+COUNTIFS($G${ph}:G{r},G{r},$F${ph}:F{r},F{r}))'
for c1,c2,label,formula,fg,bg in [(1,2,'Open follow-up tasks',f'SUM(F{ph}:F{pend})',INK,PALE),(3,4,'Require action now',f'SUM(G{ph}:G{pend})',BAD_TEXT,ROSE),(5,6,'Need follow-up',f'SUM(B{ph}:B{pend})+SUM(D{ph}:D{pend})',WARN_TEXT,YELLOW),(7,8,'Worksites with tasks',f'COUNTIF(F{ph}:F{pend-1},">0")',TEAL,MINT)]:
 merge(hub,8,c1,c2,label,fg,bg,12,False);ce=merge(hub,9,c1,c2,'='+formula,fg,bg,23,False);ce.number_format='#,##0'
hub.row_dimensions[8].height=23;hub.row_dimensions[9].height=36
merge(hub,11,1,8,'First ten by immediate tasks, then by total tasks. Red = expired documents or held time; amber = missing submissions or upcoming renewals.',INK,WHITE,11);hub.row_dimensions[11].height=24
for j,lab in enumerate(['Client / worksite','Missing time','Expired Iqamas','Renewals due','Time holds','Open tasks','Priority','Next step'],1):style(hub.cell(12,j,lab),WHITE,NAVY,11,True,'left',1,True)
hub.row_dimensions[12].height=25
for r in range(13,23):
 ix=f'$I{r}';hub[f'I{r}']=f'=IFERROR(MATCH(ROW()-12,$H${ph}:$H${pend},0),"")';hub[f'I{r}'].number_format=';;;'
 for j in range(1,7):
  col=L(j);value(hub,r,j,f'=IF({ix}="","",INDEX(${col}${ph}:${col}${pend},{ix}))',fmt=None if j==1 else '#,##0')
 value(hub,r,7,f'=IF({ix}="","",IF(C{r}+E{r}>0,"Act now","Follow up"))');value(hub,r,8,f'=IF({ix}="","",IF(C{r}>0,"Review overdue Iqamas",IF(E{r}>0,"Resolve time holds",IF(D{r}>0,"Plan renewals","Request timesheets"))))');hub.row_dimensions[r].height=23
for area in ('A13:A22','G13:G22'):
 ink_rule(hub,area,'$G13="Act now"',BAD_TEXT);ink_rule(hub,area,'$G13="Follow up"',WARN_TEXT)
merge(hub,24,1,8,'=IF(A9=0,"No workload tasks match the dashboard filters.","Showing "&COUNT(I13:I22)&" of "&COUNTIF(F2901:F2941,">0")&" groups with tasks. Counts are follow-ups, not unique employees or hours.")',MUT,PALE,11);hub.row_dimensions[24].height=24
merge(hub,26,1,8,'Counting rules',WHITE,NAVY,12,False);merge(hub,27,1,8,'One task per missing active-employee timesheet, expired Iqama, Iqama due within the configured window, or held payroll line. Past-due and upcoming expiry bands do not overlap.',INK,WHITE,11);hub['A27'].alignment=align('left',1,True);hub.row_dimensions[27].height=33
merge(hub,28,1,8,'Unlisted or missing sites are included in task totals, but not the worksite count. A person may have more than one task. The pressure view does not estimate staffing capacity or time required.',INK,PALE,11);hub['A28'].alignment=align('left',1,True);hub.row_dimensions[28].height=33
# Counts and selected-scope exception controls are distinct and explicitly labelled.
band(hub,30,'Queue controls',1,8)
for j,lab in [(1,'Audit scope'),(3,'Page · 420 per page'),(5,'Findings'),(7,'Valid rule checks')]:value(hub,31,j,lab)
value(hub,31,2,'Selected scope',True).protection=unlock;value(hub,31,4,1,True,'0').protection=unlock;dv(hub,'B31','list',arr(['Selected scope','All periods']),blank=False);dv(hub,'D31','whole',1,4,'between','Choose a queue page (420 findings per page).')
for val in list(hub.data_validations.dataValidation):
 if str(val.sqref) in ('D67','G67'):hub.data_validations.dataValidation.remove(val)
hub['D67']='=$B$31';hub['G67']='=$D$31';hub['D67'].protection=xl.styles.Protection(locked=True);hub['G67'].protection=xl.styles.Protection(locked=True);jump(hub,'A68',op_name,'A30','Queue controls')
hub['F31']='=SUM(AB2002:AB2005)';hub['H31']='=COUNTIF(D1767:D1778,"OK")&" / 12"';hub.row_dimensions[31].height=23
for lo,hi,label,loc in [(1,2,'Review findings','A69'),(3,4,'Add a movement','A955'),(5,6,'Edit settings','A1766'),(7,8,'Dashboard filters','B8')]:
 cc=merge(hub,33,lo,hi,label,TEAL,MINT,11,False);cc.hyperlink=Hyperlink(ref=cc.coordinate,location=f"'{op_name if lo!=7 else 'Dashboard'}'!{loc}",display=label)
# Helpers must remain in the same actual sheet, not a disguised extra tab.
# Use hidden row groups so no control/entry column disappears in other sections.
hub.row_dimensions.group(2001,pend,outline_level=1,hidden=True);hub.row_dimensions[2000].collapsed=True
for lo,hi in [(35,63),(933,949),(1741,1759),(1931,1999)]:hub.row_dimensions.group(lo,hi,outline_level=1,hidden=True)
hub.print_area='A1:H33';hub.page_setup.fitToHeight=1;hub.page_setup.fitToWidth=1;hub.print_options.horizontalCentered=True;hub.sheet_view.tabSelected=False
stage['Dashboard']['L6'].value='Operations control';stage['Dashboard']['L6'].hyperlink.location=f"'{op_name}'!A1";stage['Dashboard']['L6'].hyperlink.display='Operations control'
# Uniform terminology in description strings. Historical source citations are
# left intact; fix the actual new working locations rather than source labels.
for sh in stage:
 for cc in sh._cells.values():
  if cc.data_type!='s' or not isinstance(cc.value,str) or cc.hyperlink:continue
  if sh==hub and cc.row==1815:continue
  st=cc.value
  for a,b in [('in Status History','in the movement journal'),('on Audit','in the audit section'),('on Control Center','in operating settings'),('Open Audit.','Open the audit section.'),('Use Audit to','Use the audit section to'),('Salarysheet and Audit','Salarysheet and the exception queue'),('Audit →','Operations Control →'),('in Audit.','in the audit section.'),('in Audit;','in the audit section;')]:st=st.replace(a,b)
  if st!=cc.value:cc.value=st
stage['Main Dataset']['A2']='One record per IQAMA. Filter the master by client or work status. Document dated changes in Operations Control → Status history.'
# Align equivalent user-facing table columns without rewriting historical
# labels inside the received archive. Historical dates and statuses retain
# specific names to distinguish them from current master attributes.
for sh_name,table_name,edits in [(M,'EmployeeMaster',{2:'Employee name',3:'IQAMA'}),(T,'TimesheetTransactions',{3:'Employee name'}),(op_name,'AuditExceptions',{1:'Client / worksite',3:'Employee name'}),(op_name,'MovementHistory',{1:'Client / worksite',4:'Employee name'})]:
 s=stage[sh_name];tbl=s.tables[table_name];first=CellRange(tbl.ref).min_row
 for j,label in edits.items():s.cell(first,j).value=label;tbl.tableColumns[j-1].name=label
for sh_name,source_row,sample_header,width in [(M,5,357,41),(T,5,427,50),(op_name,69,572,9),(op_name,955,1377,10)]:
 for j in range(1,width+1):
  if stage[sh_name].cell(sample_header,j).value is not None:stage[sh_name].cell(sample_header,j).value=stage[sh_name].cell(source_row,j).value
stage[T]['C5'].comment=Comment('Submitted employee name; automatically compared with the employee master.','Basma / Operations');hub['A955'].comment=Comment('Client at the event date, not necessarily the employee’s current assignment.','Basma / Operations')
# Formula-built links have two parts: destination and readable row label.
# For moved targets, both must refer to the same new row.
for rr in range(70,490):
 c=hub[f'E{rr}'];c.value=re.sub(r'("History row "&\(\'Operations Control\'!AC\d+\+)5\)',r'\g<1>955)',c.value);c.value=re.sub(r'("Rule row "&\(\'Operations Control\'!AC\d+\+)6\)',r'\g<1>1766)',c.value)
merge(hub,2000,1,8,'Slicer data · grouped calculation rows for filters, charts and workload measures. Not an input area.',MUT,PALE,11);hub.row_dimensions[2000].height=26;jump(hub,'H1763',op_name,'A2000','Helper data')
# Export definitions must describe the final schema, not intermediate addresses.
for name,(sheet_name,ref,desc) in list(core_db.items()):
 if sheet_name in op_offsets:core_db[name]=(op_name,op_addr(sheet_name,ref),desc)
for n,ref in [('OperationsHome','A1:H33'),('WorkloadPressure','A12:H22'),('OperationsRules','A1766:D1778'),('OperationsHelpers',f'A2001:AC{pend}')]:
 stage.defined_names.add(DefinedName(n,attr_text=f"'{op_name}'!"+re.sub(r'([A-Z]+)(\d+)',r'$\1$\2',ref)))
# Consent-based correction overlay. Keep the entire pasted record immutable
# to the calculation: only explicitly accepted, unambiguous identity values
# feed the effective time/payroll identity. No macros or fuzzy ID inference.
from openpyxl.worksheet.table import TableColumn
ft=stage[T];fm=stage[M];fp=stage[P];fh=stage[op_name]
fix_first=870;fix_last=fix_first+(N-6);pair_col='AE';pair_range=f"'{op_name}'!${pair_col}$2006:${pair_col}$2355"
def normstr(ref):return f'TRIM(SUBSTITUTE({ref}&"",CHAR(160)," "))'
def pairkey(name,site):
 a=normstr(name);b=normstr(site);return f'LEN({a})&":"&{a}&"|"&LEN({b})&":"&{b}'
def idok(ref):return f'IFERROR(AND(LEN({ref}&"")=10,TEXT(VALUE({ref}),"0000000000")={ref}&""),FALSE)'
def mlk(col,ix):return f'INDEX({rng(M,col,E)},{ix})'
def tab_extend(sh,name,ref,headers):
 tb=sh.tables[name];tb.ref=ref;tb.autoFilter.ref=ref;first=CellRange(ref).min_row
 for j,h in headers:
  ce=sh.cell(first,j);ce.value=h;style(ce,WHITE,TEAL,10,True,'left',1,True);tb.tableColumns.append(TableColumn(id=j,name=h))
 assert len(tb.tableColumns)==CellRange(ref).max_col-CellRange(ref).min_col+1
# A literal name/client pair must be unique. Length prefixes distinguish
# field boundaries; MATCH escapes Excel wildcards. No partial-name matches.
for rr in range(6,E+1):
 name=f"'{M}'!B{rr}";site=f"'{M}'!A{rr}"
 fh[f'{pair_col}{rr+2000}']=f'=IF(OR({name}="",{site}=""),"",{pairkey(name,site)})'
# Compact, row-relative validation list: keep submitted, or the exact offered
# before/after snapshot. A generic, copied "Accept" is not valid consent.
ft['BV4']='Review and choose';ft['BW4']='Source facts are retained. Accepted identities are used by payroll.'
for r in range(6,N+1):
 h=fix_first+r-6;ID=f'D{h}';ix=f'G{h}';valid=idok(ID);mids=rng(M,'Z',E);any_input=f'BN{r}';orig_id=f'$D{fix_first}:$D{fix_last}'
 fk=f'$S{h}';name0=normstr(f'C{r}');site0=normstr(f'A{r}');cnorm=normstr(f'H{h}');snorm=normstr(f'I{h}');rawnm=f'{name0}<>{cnorm}';rawst=f'{site0}<>{snorm}'
 ft[f'D{h}']=f'=IF(D{r}="","",{normstr(f"D{r}")})'
 ft[f'S{h}']=f'=IF(OR({name0}="",{site0}=""),"",{pairkey(f"C{r}",f"A{r}")})'
 ft[f'E{h}']=f'=IF(AND({valid},COUNTIF({mids},{ID})=1),IFERROR(MATCH({ID},{mids},0),0),0)'
 ft[f'F{h}']=f'=IF(OR(BN{r}=0,D{h}<>"",{fk}="",LEN({fk})>240),0,IF(SUMPRODUCT(--({pair_range}={fk}))=1,IFERROR(MATCH({literal_criteria(fk)},{pair_range},0),0),0))'
 ft[f'G{h}']=f'=IF(E{h}>0,E{h},F{h})'
 for col,col_src in [('H','B'),('I','A'),('J','Z')]:
  ref=mlk(col_src,ix);ft[f'{col}{h}']=f'=IF({ix}=0,"",IF(ISBLANK({ref}),"",{ref}&""))'
 rawchecks=[(f'{ID}=""','ID missing'),(f'NOT({valid})','Invalid IQAMA'),(f'COUNTIFS({orig_id},{ID},$B$6:$B${N},B{r})>1','Duplicate entry'),(f'COUNTIF({mids},{ID})>1','Duplicate master ID'),(f'E{h}=0','ID not found'),(f'AND({rawnm},{rawst})','Name & site mismatch'),(f'OR(C{r}="",{rawnm})','Name mismatch'),(f'OR(A{r}="",{rawst})','Site mismatch')]
 body='"Matched"'
 for test,label in reversed(rawchecks):body=f'IF({test},"{label}",{body})'
 ft[f'C{h}']=f'=IF(BN{r}=0,"",{body})'
 # Never replace an unknown/nonblank ID. Do not guess when both supplied
 # name and site contradict the ID, or a candidate ID/month is duplicated.
 goodsite='AND('+','.join(f'I{h}<>"{st}"' for st in ['','-','Runaway','Ran away','Run away','ZZZ INACTIVE','VACATION'])+')'
 reasons=f'OR(C{h}="Name mismatch",C{h}="Site mismatch",AND(C{h}="Name & site mismatch",OR({name0}="",{site0}="")),C{h}="ID missing")'
 ft[f'K{h}']=f'=IF({ix}=0,0,--AND(BF{r}>0,{idok(f"J{h}")},COUNTIF({mids},J{h})=1,H{h}<>"",H{h}<>"-",{goodsite},{reasons},COUNTIFS($J${fix_first}:$J${fix_last},J{h},$B$6:$B${N},B{r})=1))'
 # Readable snapshot. Reserved separators are disallowed instead of silently
 # truncating/colliding values. The choice is specific to row, month, ID,
 # submitted name/site, and the exact proposed master identity.
 def before_after(label,old,new):
  show=lambda ref:f'IF({ref}="","[blank]",{ref}&"")'
  return f'"{label} "&{show(old)}&IF(EXACT({old}&"",{new}&""),""," → "&{show(new)})'
 fields=[f'D{r}',f'C{r}',f'A{r}',f'J{h}',f'H{h}',f'I{h}'];concat='&'.join(fields)
 safe='AND('+','.join(f'ISERROR(FIND("{mark}",{concat}))' for mark in ['|','→','[blank]'])+')'
 snap='&" | "&'.join(before_after(label,old,new) for label,old,new in [('ID',f'D{r}',f'J{h}'),('Name',f'C{r}',f'H{h}'),('Client',f'A{r}',f'I{h}')])
 ft[f'T{h}']=f'=IF(K{h}=0,"","Accept · R{r} | "&TEXT(B{r},"yyyy-mm-dd")&" | "&{snap})'
 ft[f'A{h}']='Keep submitted';ft[f'B{h}']=f'=IF(AND(K{h}=1,LEN(T{h})<=250,{safe}),T{h},"")'
 ft[f'L{h}']=f'=--IFERROR(AND(B{h}<>"",EXACT(BV{r},B{h}),NOT(_xlfn.ISFORMULA(BV{r}))),FALSE)'
 for col,src1,src2 in [('M','I','A'),('N','H','C'),('O','J',None)]:
  original=f'{src2}{r}&""' if src2 else f'D{h}';ft[f'{col}{h}']=f'=IF(BN{r}=0,"",IF(L{h}=1,{src1}{h},{original}))'
 def diff(label,old,new):return f'IF({normstr(old)}<>{normstr(new)},"{label}: "&IF({old}="","[missing]",{old}&"")&" → "&{new}&". ","")'
 ft[f'P{h}']='=IF(G'+str(h)+'=0,"",'+ '&'.join(diff(label,old,new) for label,old,new in [('IQAMA',f'D{r}',f'J{h}'),('Name',f'C{r}',f'H{h}'),('Client',f'A{r}',f'I{h}')])+')'
 ft[f'Q{h}']=f'=IF(BN{r}=0,"",IF(L{h}=1,"Applied",IF(AND(BV{r}<>"",BV{r}<>"Keep submitted"),"Stale",IF(B{h}<>"",IF(BV{r}="Keep submitted","Declined","Suggested"),IF(C{h}="Matched","Clear","Manual")))))'
 # Only the canonical fields change. Raw values in A:D continue to feed the
 # audit revision and the before/after note, and are never replaced by formulas.
 ft[f'BM{r}']=f'=O{h}';ft[f'AY{r}']=f'=IF({idok(f"BM{r}")},IFERROR(MATCH(BM{r},{mids},0),0),0)'
 ft[f'AZ{r}']=f'=IF(M{h}="",0,IFERROR(MATCH({literal_criteria(f"M{h}")},{rng(M,"A",C+787,793)},0),0))'
 ft[f'E{r}']=f'=IF(BN{r}=0,"",IF(AND({idok(f"BM{r}")},COUNTIFS($BM$6:$BM${N},BM{r},$B$6:$B${N},B{r})>1),"Duplicate entry",IF(Q{h}="Stale","Review changed suggestion",IF(L{h}=1,"Matched",C{h}))))'
 for col in ['BE','BP','BT']:
  ft[f'{col}{r}']=re.sub(r'(?<![A-Z!])A'+str(r)+r'(?!\d)',f'M{h}',ft[f'{col}{r}'].value)
 for col,eff in [('A','M'),('C','N'),('D','O')]:fp[f'{col}{r}']=f'=IF(\'{T}\'!BN{r}=0,"",\'{T}\'!{eff}{h})'
 for j,eff in [(1,'M'),(3,'N'),(4,'O')]:
  row=435+r-6;ref=f'INDEX(${eff}${fix_first}:${eff}${fix_last},$BU{row})';ft.cell(row,j).value=f'=IF($BU{row}="","",{ref})'
 # Explanations stay useful even when automatic repair is unavailable, and
 # a valid identity correction never suppresses a second attendance/pay rule.
 advise=f'IFERROR(VLOOKUP(BR{r},RepairRules,2,FALSE),"Verify the source record in Operations Control.")'
 ft[f'R{h}']=f'=IF(OR(BN{r}=0,BR{r}="Checks passed"),"",BR{r}&": "&IF(BR{r}="Incomplete days",AP{r}&" of "&BF{r}&" days logged. ","")&{advise})'
 other=f'IF(AND(BR{r}<>"Checks passed",BR{r}<>C{h},BR{r}<>"Review changed suggestion")," Also: "&R{h},"")'
 manual=f'IF(C{h}="ID missing",IF(G{h}=0,"No unique exact name/client match. Verify the IQAMA document.","Verify this ID and check for a repeated ID/month before entering it."),IF(OR(C{h}="Invalid IQAMA",C{h}="ID not found"),"Submitted ID: "&D{r}&". Verify the document; do not substitute another ID.",IF(P{h}<>"",P{h}&"Check supporting evidence before editing.",IFERROR(VLOOKUP(C{h},RepairRules,2,FALSE),"Verify against the original source."))))'
 ft[f'BW{r}']=f'=IF(BN{r}=0,"",IF(Q{h}="Applied","APPLIED · "&P{h}&IF(BR{r}="Checks passed","All other checks passed.","Still held: "&R{h}),IF(Q{h}="Stale","RECHECK · The saved decision no longer matches this submission or master. Clear it, then review the current facts. "&P{h},IF(OR(Q{h}="Suggested",Q{h}="Declined"),IF(Q{h}="Declined","NOT APPLIED · ","SUGGESTED · ")&P{h}&"Select the row-specific Accept option in Decision to use the master values."&{other},IF(C{h}<>"Matched","VERIFY · "&C{h}&". "&{manual}&{other},IF(BR{r}="Checks passed","OK · No discrepancy.","VERIFY · "&R{h}))))))'
 ce=value(ft,r,74,None,True,'@');ce.protection=unlock;ce.alignment=Alignment(horizontal='left',vertical='center',indent=1)
 # Two cells in the same sheet are a native, Excel-compatible list source.
 # Use a separate rule per row: no INDIRECT/OFFSET and no cross-sheet DV links.
 dv(ft,f'BV{r}','list',f'$A${h}:$B${h}',msg='Read the rightmost column. Select the matching Accept option or Keep submitted. Delete the selection to revoke it.')
 ce=ft[f'BW{r}'];style(ce,INK,PALE,11,False,'left',1,True)
 jr=435+r-6;ft[f'J{jr}']=f'=IF($BU{jr}="","",INDEX($BW$6:$BW${N},$BU{jr}))'
 # Existing signed reviews must be redone after an approval is changed. The
 # payroll lock still depends on all of the original rule and review checks.
 ar=r+64;fh[f'Y{ar}']=fh[f'Y{ar}'].value[:-1]+f'&"|decision:"&\'{T}\'!BV{r})';fh[f'U{ar}']=f'=IF(\'{T}\'!BN{r}=0,"",\'{T}\'!BW{r})'
for rr in range(1103,1143):
 for co in ['D','L']:
  ce=stage[M][f'{co}{rr}'];ce.value=ce.value.replace(f"'{T}'!$A$6:$A${N}",f"'{T}'!$M${fix_first}:$M${fix_last}")
# Discrepancy is the last *visible* column. BT:BU are pre-existing list
# indices, so give them real unique names instead of blank table headers.
tab_extend(ft,'TimesheetTransactions',f'A5:BW{N}',[(72,'Quick filter order'),(73,'Review row index'),(74,'Correction decision'),(75,'Discrepancy / recommendation')]);ft.column_dimensions['BV'].width=28;ft.column_dimensions['BW'].width=78
ft['BV5'].comment=Comment('Acceptance is a literal snapshot of this row’s month, original identity and proposed master values. Select it from the list; a formula or an old decision is not approval.','Basma / Operations')
ft['BW5'].comment=Comment('Amber: proposed correction; green: accepted or clear; red: verify or review changed evidence. Submitted values are kept unchanged. Select the cell for the complete note.','Basma / Operations')
ft['J434']='Discrepancy / recommendation';style(ft['J434'],WHITE,TEAL,10,True,'left',1,True)
ft['A431']='Filter by client, result and month. The rightmost note explains the issue; open its time row, then choose Review fixes. Accepted values appear here and in payroll.'
jump(ft,'G3',T,'BV5','Review fixes →');ft['A2']='Paste values in A:D and H:AL. E:G checks identity. Use Review fixes to accept a supported master correction; A:D always retains the submitted record.'
ft.row_dimensions.group(fix_first,fix_last,hidden=True,outline_level=1)
# Extend the existing one-row reference without creating a live suggestion
# or decision. The archive sample already matches the master.
for sh,cols in [(ft,[74,75]),(fp,[80])]:
 for j in cols:
  style(sh.cell(427,j),TEAL,PALE,10,True,'left',1,True)
  sh.cell(427,j).value='Correction decision' if j==74 else 'Discrepancy / recommendation'
ft['BW428']='OK · Reference row; the archived identity matches the master.'
fp['CB428']="='Timesheet_Input'!BW428"
# The monthly view has narrow weekday columns; span its explanation across
# those columns rather than widening the underlying daily entry grid.
for rr in range(434,855):
 ft.merge_cells(start_row=rr,start_column=10,end_row=rr,end_column=15);ft[f'J{rr}'].alignment=align('left',1,True)
 if rr>434:ft.row_dimensions[rr].height=35
# Present corrections alongside calculated results and the fixed time row.
# A remaining hold must stay visibly amber/red even if one identity fix applied.
for r in range(6,N+1):
 fp[f'CB{r}']=f'=IF(\'{T}\'!BN{r}=0,"",\'{T}\'!BW{r})';style(fp[f'CB{r}'],INK,PALE,11,False,'left',1,True)
tab_extend(fp,'PayrollResults',f'A5:CB{N}',[(80,'Discrepancy / recommendation')]);fp.column_dimensions['CB'].width=68
for sh,area,col,release in [(ft,f'BW6:BW{N}','BW6','BQ6'),(ft,'J435:J854','J435','F435'),(fp,f'CB6:CB{N}','CB6','O6')]:
 # Make an applied-but-not-released correction unambiguously different from
 # a full pass, even on clients which ignore lower-priority differential rules.
 if area in [str(obj.sqref) for obj in sh.conditional_formatting]:del sh.conditional_formatting[area]
 ink_rule(sh,area,f'OR(LEFT({col},6)="VERIFY",LEFT({col},7)="RECHECK",AND(LEFT({col},7)="APPLIED",{release}="Hold"))',BAD_TEXT)
 ink_rule(sh,area,f'OR(LEFT({col},9)="SUGGESTED",LEFT({col},11)="NOT APPLIED")',WARN_TEXT)
 ink_rule(sh,area,f'OR(AND(LEFT({col},7)="APPLIED",{release}="Ready"),LEFT({col},2)="OK")',OK_TEXT)
core_db['PayrollResults']=(P,f'A5:CB{N}','Effective client/name/ID, pay amounts and release checks; the last column preserves the before/after explanation.')
# Provide actionable notes even for exceptions for which no safe edit can
# be derived. The populated 17 salary findings stay visible pending evidence.
master_advice={'Employee name missing':'Enter the documented full name; do not infer a name from the ID alone.','Invalid master IQAMA':'Verify all ten digits against the document.','Duplicate master ID':'Resolve the duplicated IQAMA against the original employee files.','Unrecognized status':'Choose the documented work status from the list.','Client missing':'Confirm the current assignment and enter its client.','Check master salary':'Verify the signed contract; enter the documented positive SAR salary in column P. A rate-card estimate is not proof.','Missing IQAMA date':'Enter the expiry date from the current IQAMA.','Missing insurance date':'Enter the documented insurance expiry.'}
for i,(key,tip) in enumerate(master_advice.items(),3001):fh[f'AA{i}']=key;fh[f'AB{i}']=tip
fh.row_dimensions.group(3001,3000+len(master_advice),outline_level=1,hidden=True)
for r in range(6,E+1):
 actual=f'IF(AU{r}="Check master salary"," Submitted salary: "&IF(P{r}="","[missing]",P{r}&"")&".",IF(OR(AU{r}="Invalid master IQAMA",AU{r}="Duplicate master ID")," Submitted ID: "&C{r}&".",""))'
 fm[f'AW{r}']=f'=IF(OR(AU{r}="",AU{r}="OK"),"","VERIFY · "&AU{r}&"."&{actual}&" "&IFERROR(VLOOKUP(AU{r},\'{op_name}\'!$AA$3001:$AB$3008,2,FALSE),"Check the employee evidence."))';style(fm[f'AW{r}'],INK,PALE,11,False,'left',1,True)
tab_extend(fm,'EmployeeMaster',f'A5:AW{E}',[(48,'Master exception order'),(49,'Discrepancy / recommendation')]);fm.column_dimensions['AW'].width=68;ink_rule(fm,f'AW6:AW{E}','AW6<>""',BAD_TEXT)
# The exception queue has a moving order; NEVER place editable permissions
# beside it. Its rightmost cell links to the stable source decision or input.
event_tips={'Event ID missing':'Enter a unique event reference from the documented change.','Duplicate event ID':'Compare the two source events; retain separate unique references only for genuinely different events.','Invalid IQAMA':'Verify the full ten-digit ID from the employee document.','Check employee ID':'Verify the IQAMA against one and only one employee in Main Dataset.','Check effective date':'Enter the actual documented event date; do not invent a date or enter a future event.','Check event type':'Select the event type supported by the source record.','Evidence missing':'Enter the actual document reference supporting this event.','Check transfer sites':'Confirm distinct previous and new client assignments.','Check status change':'Enter distinct, documented From and To employment statuses.','Check exit status':'Select the documented exit status: Terminated, Left or Resigned.','Check case status':'Select the documented case state: Reported, Under review or Resolved.','Check correction':'Record different prior and corrected values using the evidence.','Duplicate event':'Compare the dates, employee and movement details; consolidate only a confirmed duplicate.'}
for i,(key,tip) in enumerate(event_tips.items(),3011):fh[f'AA{i}']=key;fh[f'AB{i}']=tip
fh.row_dimensions.group(3011,3010+len(event_tips),outline_level=1,hidden=True)
fh['J69']='Recommendation / action';fh.tables['AuditExceptions'].tableColumns[9].name='Recommendation / action';fh.column_dimensions['J'].width=44
for r in range(70,490):
 h=r+2391;kind=f'$AB${h}';ix=f'$AC${h}';idx=lambda sh,col,end:f'INDEX({rng(sh,col,end)},{ix})'
 time_note=idx(T,'BW',N);master_note=idx(M,'AW',E);history_note=f'INDEX($J$956:$J$1375,{ix})';rule_note=f'INDEX($C$1767:$C$1778,{ix})'
 f=f'IF({kind}="Time",HYPERLINK("#\'{T}\'!BV"&({ix}+5),{time_note}),IF({kind}="Employee",HYPERLINK("#\'{M}\'!AW"&({ix}+5),{master_note}),IF({kind}="History",HYPERLINK("#\'{op_name}\'!J"&({ix}+955),"VERIFY · "&{history_note}&". "&IFERROR(VLOOKUP({history_note},$AA$3011:$AB$3023,2,FALSE),"Confirm the event evidence.")),HYPERLINK("#\'{op_name}\'!B"&({ix}+1766),"VERIFY · "&F{r}&". "&{rule_note}))))'
 fh[f'J{r}']=f'=IF({ix}=0,"",{f})';fh[f'J{r}'].alignment=align('left',1,True)
ink_rule(fh,'J70:J489','OR(LEFT(J70,6)="VERIFY",LEFT(J70,7)="RECHECK")',BAD_TEXT);ink_rule(fh,'J70:J489','OR(LEFT(J70,9)="SUGGESTED",LEFT(J70,11)="NOT APPLIED")',WARN_TEXT)
fh['A1859']='Review changed suggestion';merge(fh,1859,2,8,'Clear the outdated decision in Timesheet_Input BV, then check the new recommendation and select the matching offer.',INK,PALE,11)
stage.defined_names.add(DefinedName('RepairRules',attr_text=f"'{op_name}'!$A$1838:$B$1859"))
fh['A1930']='Accepted corrections';merge(fh,1930,2,8,'Timesheet_Input → Review fixes. BV is the decision; BW shows submitted → proposed values. Accept only after checking the source. Clear the decision to revert. Non-unique or unsupported corrections require evidence. Payroll and reports use effective values; A:D retain the paste.',INK,PALE,11);fh['B1930'].alignment=align('left',1,True);fh.row_dimensions[1930].height=65
core_db['TimesheetTransactions']=(T,f'A5:BW{N}','One IQAMA/month; raw paste retained. A row-specific accepted master correction drives effective identity and client.')
core_db['EmployeeMaster']=(M,f'A5:AW{E}','Single authoritative employee master; the final column explains evidence-required discrepancies.')

# Final calculation pass uses only the consolidated layout. Check *every*
# historic data/formula result and all six tables rather than just headline KPIs.
with tempfile.TemporaryDirectory(prefix='basma-ops-') as work:
 work=Path(work);src_stage=work/'operations.xlsx';stage.save(src_stage);calculated=recalc(src_stage,work/'computed');final_cache=xl.load_workbook(calculated,data_only=True)
 errors=[(sh.title,c.coordinate,c.value)for sh in final_cache for c in sh._cells.values() if c.data_type=='e'];assert not errors,errors[:10]
 for chartsh in ('Dashboard',P):
  for chart in stage[chartsh]._charts:
   for ser in chart.series:assert f"'{op_name}'!" in ser.val.numRef.f and ser.val.numRef.numCache is not None
 assert final_cache.sheetnames==['Dashboard',M,T,P,op_name]
 # Changing tabs and row addresses cannot change operational amounts. Confirm
 # all non-presentation formulas still return their old results. Exclusions
 # are the captions of row links, which now carry their new physical row.
 exceptions=[]
 for prev in base_view:
  target=final_cache[op_name if prev.title in op_offsets else prev.title];dy0=op_offsets.get(prev.title,0)
  for cc in prev._cells.values():
   if cc.data_type not in ('n','d','b') or cc.value is None:continue
   before=cc.value;after=target.cell(cc.row+dy0,cc.column).value
   if before!=after and not(isinstance(before,(int,float)) and isinstance(after,(int,float)) and abs(before-after)<.00001):exceptions.append((prev.title,cc.coordinate,before,after))
 assert not exceptions,exceptions[:10]
 oc=final_cache[op_name];assert [oc[f'{x}9'].value for x in 'ACE']==[final_cache['Dashboard']['F111'].value+final_cache['Dashboard']['J108'].value+final_cache['Dashboard']['N108'].value+final_cache['Dashboard']['F20'].value,final_cache['Dashboard']['J108'].value+final_cache['Dashboard']['F20'].value,final_cache['Dashboard']['F111'].value+final_cache['Dashboard']['N108'].value]
 assert oc['F31'].value==17 and oc['H31'].value=='12 / 12' and oc['E1383'].value==2
 assert len(set(oc[f'A{r}'].value for r in range(13,23) if oc[f'A{r}'].value))==sum(oc[f'A{r}'].value is not None for r in range(13,23))
 assert all(c.value in ('Act now','Follow up') for c in [oc[f'G{r}']for r in range(13,23)] if c.value is not None)
 # Commit the consolidated package only after the relocated engine passes.
 attach_cache(src_stage,calculated,out);check=xl.load_workbook(out);final_cache=xl.load_workbook(out,data_only=True)
for sh in check:
 for cc in sh._cells.values():
  if cc.data_type=='f' and isinstance(cc.value,str):assert not any("'"+name+"'!" in cc.value for name in old_tabs),(sh.title,cc.coordinate,cc.value)
  if cc.hyperlink and cc.hyperlink.location:assert cc.hyperlink.location.rsplit('!',1)[0].strip("'") in check.sheetnames,(sh.title,cc.coordinate,cc.hyperlink.location)
for n in check.defined_names.values():assert not any("'"+name+"'!" in n.attr_text for name in old_tabs)
assert all(check[op_name].cell(rr+1760,2).protection.locked is False for rr in range(7,19)) and check[op_name]['V70'].protection.locked is False and check[op_name]['E1026'].protection.locked is False
assert sum(s.sheet_state=='visible' for s in check)==5 and len(check.sheetnames)==5 and check[op_name].row_dimensions[2001].hidden
print('PASS: Audit, Status History, Control Center and calculation helpers now share Operations Control. Five visible tabs, no hidden worksheets; workload totals reconcile to the dataset and time register.')

# Verify the *relocated* controls, including downstream payroll and current
# data validations. Run mutations only in isolated files; never seed live input.
if os.environ.get('BASMA_OPS_VERIFY')=='1':
 with tempfile.TemporaryDirectory(prefix='basma-hub-test-') as qa:
  qa=Path(qa);trial=xl.load_workbook(out);orig=xl.load_workbook(out,data_only=True);H=trial[op_name];ts=trial[T];dash=trial['Dashboard'];master0=trial[M]
  def test_hub(phase):
   dr=qa/phase;dr.mkdir();f=dr/'test.xlsx';trial.save(f);recalced=recalc(f,dr/'out');vv=xl.load_workbook(recalced,data_only=True)
   assert not [(s.title,c.coordinate,c.value) for s in vv for c in s._cells.values() if c.data_type=='e'],phase
   ho=vv[op_name];tasks=sum(ho.cell(p,k).value or 0 for p in range(ph,pend+1) for k in (2,3,4,5));assert tasks==ho['A9'].value==ho['C9'].value+ho['E9'].value
   assert ho['F31'].value==sum(ho[f'AB{x}'].value for x in range(2002,2006));return vv
  for r,ar in [(6,497),(7,498),(8,499)]:
   for j,k in [(1,3),(2,2),(3,4),(4,5),*[(i+8,i+6) for i in range(30)]]:ts.cell(r,j).value=orig[op_name].cell(ar,k).value
  ts['C7']='QA name mismatch';ts['A8']='QA UNLISTED SITE';H['B31']='All periods';H['B1771']='Invalid policy';H['B1770']=dt.datetime(2026,11,19)
  for j,val in {1:'SHETA & SAIF',2:dt.datetime(2026,9,5),3:orig[op_name]['E497'].value,5:'QA-MOVE-01',6:'Status change',7:'Active',8:'On Leave',9:'Verified test evidence'}.items():H.cell(1026,j).value=val
  for j,val in {1:'SHETA & SAIF',2:dt.datetime(2026,9,6),3:'0000000000',5:'QA-MOVE-02',6:'Joining',9:'Test missing master'}.items():H.cell(1027,j).value=val
  vv=test_hub('invalid-policy');ho=vv[op_name];assert vv[P]['P6'].value=='Check payroll settings' and vv[P]['N6'].value==0 and ho['D1771'].value=='Check rule' and ho['H31'].value=='11 / 12'
  assert ho['J1026'].value=='Recorded' and ho['J1027'].value=='Check employee ID';assert ho['AB2004'].value==1 and ho['AB2005'].value==1;assert ho['AB2002'].value==3
  assert ho['E2941'].value==1 and ho['G2941'].value>=1
  # Date/rate validation, reviewed signatures and pay must still respond
  # to the merged editing cells and the preserved table identifiers.
  H['B1771']=7;H['V71']='TEST REVIEWER';H['W71']=dt.datetime.today().replace(hour=0,minute=0,second=0,microsecond=0);H['X71']='Test: compared submitted name with the source.';ts['C7']=orig[op_name]['D498'].value
  dash['F9']='SHETA & SAIF';H['B31']='Selected scope';vv=test_hub('correct-source');ho=vv[op_name]
  assert vv[P]['O6'].value=='Ready' and vv[P]['N6'].value==1200 and vv[P]['P7'].value=='Document correction';assert ho['G9'].value==1 and ho['A13'].value=='SHETA & SAIF' and ho['A14'].value is None;assert ho['F31'].value>=1
  H['Z71']=ho['Y71'].value;vv=test_hub('signed');assert vv[P]['N7'].value==1200 and vv[op_name]['S71'].value=='Remade' and vv['Dashboard']['S6'].value==2400
  ts['H7']='A';H['B31']='All periods';H['D31']=2
  # Deliberately overflow one page. All-period navigation and its row IDs
  # must span 420 lines and still point at the right journal/register rows.
  for rr in range(1026,1376):H[f'E{rr}']='QA DUPLICATE'
  for rr in range(6,106):master0[f'P{rr}']=None
  vv=test_hub('pagination');ho=vv[op_name];assert ho['F31'].value>420;assert sum(bool(ho[f'F{rr}'].value) for rr in range(70,490))==ho['F31'].value-420
  assert ho['D67'].value=='All periods' and ho['G67'].value==2 and ho['E1383'].value==2;assert vv[P]['N7'].value==0 and ho['AD71'].value==0
  print('PASS consolidated controls: client scope, workload totals, unmatched-site pressure, events, invalid settings, payroll holds, signed corrections, changed signatures and second-page exceptions.')

if os.environ.get('BASMA_CORRECTIONS_VERIFY')=='1':
 w=xl.load_workbook(out);base=final_cache;T='Timesheet_Input';P='Salarysheet';H=op_name;M='Main Dataset';t=w[T];h=w[H];m=w[M];d=w['Dashboard']
 originals={}
 for r in range(6,18):
  ar=497+r-6
  for j,k in [(1,3),(2,2),(3,4),(4,5),*[(8+i,6+i) for i in range(30)]]:t.cell(r,j).value=base[H].cell(ar,k).value
  originals[r]=[t[f'{c}{r}'].value for c in 'ABCD'];assert originals[r][0]=='SHETA & SAIF'
 t['C6']=None;t['A7']='PLASTIC';t['D8']=None;t['C9']='TEST TYPO';t['H9']='X';t['A10']='PLASTIC';t['C10']='WRONG PERSON';t['D11']='123*';t['D12']='1234567890';t['A13']=None;t['C13']=None;t['D15']=t['D14'].value;t['C16']=None;t['D17']=None
 # Ambiguous name + site cannot become an identity proposal.
 master_r={str(m[f'C{r}'].value):r for r in range(6,209)};m['A300']=originals[17][0];m['B300']=originals[17][2];m['C300']='9999997777';m['J300']='Active'
 h['B31']='All periods'
 with tempfile.TemporaryDirectory(prefix='approved-fixes-') as temp:
  root=Path(temp)
  def run(st):
   dr=root/st;dr.mkdir();f=dr/'qa.xlsx';o=dr/'out';o.mkdir();w.save(f);res=subprocess.run(['soffice','-env:UserInstallation='+(dr/'profile').as_uri(),'--headless','--convert-to','xlsx','--outdir',str(o),str(f)],capture_output=True,text=True,timeout=140);assert (o/f.name).is_file(),res.stderr
   v=xl.load_workbook(o/f.name,data_only=True);errors=[(s.title,c.coordinate,c.value) for s in v for c in s._cells.values() if c.data_type=='e'];assert not errors,errors[:20];return v
  v=run('proposals');print('Correction tests: pending suggestions checked.',flush=True)
  offers={r:v[T][f'B{r+864}'].value for r in range(6,18)}
  assert all(offers[r] for r in [6,7,8,9,13,16]);assert not any(offers[r] for r in [10,11,12,14,15,17]);assert all(v[P][f'N{r}'].value==0 for r in range(6,18));assert v[T]['BW6'].value.startswith('SUGGESTED') and v[H]['J70'].value==v[T]['BW6'].value
  for r in [6,7,8,9,13]:t[f'BV{r}']=offers[r]
  t['BV16']='Keep submitted';v=run('accepted');print('APPLIED',[(r,v[T][f'E{r}'].value,v[P][f'P{r}'].value,v[P][f'N{r}'].value)for r in [6,7,8,9,13,16]],flush=True)
  for r in [6,7,8,13]:assert v[P][f'O{r}'].value=='Ready' and v[P][f'N{r}'].value==1200 and v[P][f'C{r}'].value==originals[r][2] and v[P][f'D{r}'].value==originals[r][3] and v[P][f'A{r}'].value==originals[r][0],(r,v[P][f'P{r}'].value)
  assert v[P]['P9'].value=='Invalid day code' and v[P]['N9'].value==0 and 'Still held' in v[T]['BW9'].value;assert v[T]['BW16'].value.startswith('NOT APPLIED');assert v[T]['A7'].value=='PLASTIC' and v[T]['C6'].value is None and v[T]['D8'].value is None
  assert v['Dashboard']['B20'].value==12 and v['Dashboard']['S6'].value==4800;assert v[P]['AA7'].value==46.15 and v[P]['CB7'].value==v[T]['BW7'].value
  # Accepted identity repeated elsewhere and a stale/copied/generic decision all
  # stay blocked. Clearing a decision restores the original facts immediately.
  d['F9']='SHETA & SAIF';t['B432']='SHETA & SAIF';t['D432']='Ready';t['C6']='Other raw input';t['BV7']=None;t['B8']=dt.datetime(2026,10,1);t['BV16']='Accept';t['BV11']=offers[6];m[f'B{master_r[originals[13][3]]}']='New master spelling';v=run('stale');
  assert v[T]['E6'].value==v[T]['E8'].value==v[T]['E13'].value=='Review changed suggestion';assert v[P]['N6'].value==v[P]['N7'].value==v[P]['N8'].value==v[P]['N13'].value==0;assert v[T]['E7'].value=='Site mismatch' and v[T]['BW13'].value.startswith('RECHECK');assert v[P]['N16'].value==v[P]['N11'].value==0
  print('PASS: suggestions, literal informed acceptance, reversal, stale decisions, rejection, ambiguous/malformed/unknown IDs, duplicate holds and independent attendance checks.')

# Deliver an inspectable ZIP in addition to the same original-filename XLSX.
# The archive contains no extra copies of raw employee data or build fixtures.
import hashlib,json
package=ROOT/'Basma_HR_Database.zip'
review_guide='''BASMA | HR DATABASE REVIEW
=========================
Extract Basma_Workforce_and_Payroll.xlsx to a restricted folder and open it in Excel.
The XLSX inside this ZIP is byte-for-byte identical to the standalone workbook.
No external connections, macros, add-ins or server are needed. Use automatic calculation.
This ZIP is not encrypted: it contains personal identifiers, bank details and payroll data.

DAILY WORK
1. Start on Dashboard. The leading cards show active staff, expired Iqamas, renewals
   due in 0–15 days, and blocked time entries. Filter the renewal worklist to find
   exact employees. Roster/document checks use the current master and compliance date;
   the payroll month selects transactions, not a reconstructed historic roster.
2. Main Dataset is the current employee master. Keep one row per 10-digit IQAMA.
   Do not duplicate a person to record a transfer. Do not invent missing pay or dates.
3. In Operations Control > Status history, add documented movements, one event per row with a stable,
   unique Event ID. Then update the current master. The supplied source has 70 real
   joining dates. No past site, departure or case date was invented. The employee
   name displayed with an event is a current lookup; recorded dates/from/to/evidence
   are stored values. Correct by adding a new referenced event; this is not an
   automatic or tamper-proof edit log. Keep dated copies under access control.
4. Timesheet_Input starts empty. Paste values only into A:D (client, month start,
   name, IQAMA), and H:AL (day marks). Use the monthly review and native filters.
   Review fixes opens the right-edge Decision / Discrepancy fields. Choose the
   exact row-specific Accept option to apply a unique master match, or Keep
   submitted. Source A:D is preserved; corrected values flow to the client
   filter, payroll and reports. Clearing the choice revokes it. Changed identity,
   month or master details invalidate it. No wages, dates or day marks are guessed.
   One employee/month is payable only once. A zero day and a blank day differ.
5. Salarysheet is calculated from the same time row. Gross minus deductions equals
   estimated net; ready plus held equals estimated net. Ready is not authorization.
6. Operations Control > Audit & rework lists current exceptions and links to the source. Master, time, movement
   and configuration checks share one queue. Select All periods for a full review;
   420 findings are shown per page. After a correction, the next issue may appear.
   Payroll issues are listed first. Master or history issues do not silently change
   pay. Invalid pay-cycle, absence-share or currency settings block release.

SIX CORE TABLES
EmployeeMaster        Main Dataset      IQAMA; current facts and live checks.
TimesheetTransactions Timesheet_Input   Month-start + IQAMA; live attendance inputs.
PayrollResults        Salarysheet       Same time row/employee/month; calculated output.
MovementHistory       Operations Control   Event ID + IQAMA; dated evidence.
RulesConfig           Operations Control   Unique setting; value and validation.
AuditExceptions       Operations Control   Source record + primary finding.
Use Excel's Name Box to select a table by name. Filter by a table's header arrows.

WHAT TO RETAIN / REVIEW
Keep the five visible tabs. Operations Control combines the exception queue,
movement log, rules/config and hidden helper rows for all filters and charts.
Its workload view ranks sites by immediate items, then total tasks. Counts
are follow-ups, not hours or unique employees. Use the Dashboard filters
for workforce scope; use the queue controls for all-period exception review. The supporting ContractMatrix on Main Dataset
is required by payroll: it contains client-specific rates, shift lengths, absence
policies and OT rules. ClaimRegister feeds expense reporting; WorksiteSummary and
ClientStatement are calculated views, not competing sources of employee or pay data.
Review optional report sections on a COPY before deleting them. Hiding a section
is safer than deleting it. Samples are outside all live register ranges. The 73
original September submissions are preserved in Operations Control > Received archive, not in payroll.
The history section retains the original linked case view, including the two
source-marked runaway cases; ordinary departures and absences are not cases.

CAPACITY / LIMITS
Prepared capacity: 350 employees, 420 monthly time rows and payroll results,
420 event records, 40 client rules and 250 claims. Use prepared rows only. Do
not insert/sort a single operational register independently: time rows, signed
reviews and payroll are paired. Use filters instead. To close a payroll month,
keep a dated approved copy; changing the master/rates changes live estimates.
Protection prevents accidental edits, not unauthorized access. The workbook
is not a substitute for legal or contractual payroll verification.

REVIEW AT DELIVERY
203 source employees; 70 documented joins; zero live timesheets. Seventeen
nonnumeric or zero salaries were flagged, not guessed or overwritten. The workbook
was recalculated and its formula caches, table ranges, native filters, chart links,
document-expiry windows, payroll checks and preserved source evidence were checked.
See database_manifest.json for exact core table ranges, dependencies and file hash.
'''
manifest={'workbook':out.name,'sha256':hashlib.sha256(out.read_bytes()).hexdigest(),'source_sha256':hashlib.sha256(source_file_path.read_bytes()).hexdigest(),'core_tables':[{'name':n,'sheet':sh,'range':ref,'purpose':desc} for n,(sh,ref,desc) in core_db.items()],'linked_support_tables':['ContractMatrix','ClaimRegister','WorksiteSummary','ClientStatement'],'helper_section':{'sheet':op_name,'range':f'A2001:AC{pend}','default_visibility':'grouped and hidden rows'},'workload_measure':'Open follow-up tasks: missing active-staff timesheets + expired Iqamas + Iqamas inside the configured warning window + held payroll lines. Not hours or unique people.','source_counts':{'employees':len(ep),'dated_join_events':len(join_events),'archived_timesheet_rows':len(ids),'live_timesheet_rows':0},'limits':{'employees':E-5,'time_rows':N-5,'payroll_rows':N-5,'history_events':hist_end-5,'exceptions_per_page':N-5},'security':'Confidential; ZIP is not encrypted.'}
for n,(sh,ref,desc) in core_db.items():
 assert check[sh].tables[n].ref==ref and check[sh].tables[n].autoFilter.ref==ref
assert sum(n in sh.tables for sh in check for n in core_db)==6 and not check._external_links
with zipfile.ZipFile(package,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
 z.write(out,out.name);z.writestr('REVIEW_GUIDE.txt',review_guide);z.writestr('database_manifest.json',json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
with zipfile.ZipFile(package) as z:
 assert z.testzip() is None and z.namelist()==[out.name,'REVIEW_GUIDE.txt','database_manifest.json'];assert z.read(out.name)==out.read_bytes();assert json.loads(z.read('database_manifest.json'))['sha256']==manifest['sha256']
print(f'Verified ZIP: {package.name}; one workbook, six core named tables, review guide and dependency manifest.')
