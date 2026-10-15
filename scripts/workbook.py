"""Controlled template writes and offline structural verification."""
from collections import defaultdict
from copy import copy
from datetime import datetime, timezone
import io
import json
import os
from pathlib import Path
import re
import tempfile
import math
import unicodedata
from urllib.parse import urlsplit

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.hyperlink import Hyperlink

from inputs import canonical, check_id, norm, read_run, sha


def link_signature(c):
    h = c.hyperlink
    return (h.target,h.location,h.display,h.tooltip) if h else None


def cell_signature(c):
    # XLSX saves can deduplicate equivalent style-table indices; compare effective styles.
    alignment=copy(c.alignment)
    for flag in ['wrapText','shrinkToFit','justifyLastLine']:
        setattr(alignment,flag,bool(getattr(alignment,flag)))
    styles=tuple(str(v) for v in (c.font,c.fill,c.border,alignment,c.protection,c.number_format))
    value=format(c.value,'.14g') if isinstance(c.value,(int,float)) and not isinstance(c.value,bool) else c.value
    return (value,c.data_type,styles,link_signature(c),str(c.comment))


def baseline(w):
    return {s.title:{'cells':{c.coordinate:cell_signature(c) for row in s for c in row if c.value is not None or c.has_style or c.hyperlink},'merged':set(map(str,s.merged_cells.ranges)),'validation':str(s.data_validations),'images':len(s._images),'charts':len(s._charts)} for s in w}


def diagnostics(w):
    issues=[]
    for s in w:
        for row in s:
            for c in row:
                if c.data_type=='f':
                    if '#REF!' in c.value:
                        issues.append(f'{s.title}!{c.coordinate}: inherited #REF! formula')
                    for name in re.findall(r"'((?:[^']|'')+)'!",c.value):
                        if name.replace("''", "'") not in w.sheetnames:
                            issues.append(f'{s.title}!{c.coordinate}: formula references missing sheet {name}')
                    if '__xludf.DUMMYFUNCTION' in c.value:
                        issues.append(f'{s.title}!{c.coordinate}: Google-export placeholder formula, native recalculation unverified')
                if c.hyperlink and c.hyperlink.location and c.hyperlink.location.startswith(('heading=','gid=')):
                    issues.append(f'{s.title}!{c.coordinate}: inherited Google anchor {c.hyperlink.location}, not an Excel cell target')
    return issues


def internal_target(location):
    location=location.lstrip('#')
    if location.startswith(('heading=','gid=')):
        return None
    m=re.fullmatch(r"'((?:[^']|'')+)'!\$?([A-Z]+)\$?([1-9]\d*)", location)
    if not m:
        m=re.fullmatch(r'([^!]+)!\$?([A-Z]+)\$?([1-9]\d*)',location)
    if not m:
        raise ValueError(f'invalid internal link target: {location}')
    name=m[1].replace("''", "'")
    cell=m[2]+m[3]
    if openpyxl.utils.column_index_from_string(m[2])>16384 or int(m[3])>1048576:
        raise ValueError(f'out of bounds link target: {location}')
    return name,cell


def validate_links(w):
    count=0
    for s in w:
        for row in s:
            for c in row:
                locations=[]
                if c.hyperlink:
                    if c.hyperlink.location:
                        locations.append(c.hyperlink.location)
                    elif c.hyperlink.target and c.hyperlink.target.startswith('#'):
                        locations.append(c.hyperlink.target)
                if c.data_type=='f':
                    m=re.match(r'=HYPERLINK\("(#[^"]+)"',c.value,re.I)
                    if m:locations.append(m[1])
                for location in locations:
                    target=internal_target(location)
                    if target:
                        name,cell=target
                        if name not in w.sheetnames:
                            raise ValueError(f'internal link references missing sheet: {s.title}!{c.coordinate} -> {name}')
                        count+=1
    return count


def put(ws,row,col,value,touched,url=False):
    c=ws.cell(row,col)
    if c.data_type=='f':
        raise ValueError(f'refusing to overwrite formula: {ws.title}!{c.coordinate}')
    if c.hyperlink:
        raise ValueError(f'refusing to overwrite hyperlink data cell: {ws.title}!{c.coordinate}')
    c.value=value
    # Untrusted strings beginning '=' are literal report text, never formulas.
    if isinstance(value,str):c.data_type='s'
    touched.add((ws.title,c.coordinate))
    if url and isinstance(value,str) and value.strip().startswith(('http://','https://')):
        c.value=value.strip()
        c.hyperlink=c.value
        c.font=Font(name=c.font.name or 'Arial',size=c.font.sz or 10,color='0563C1',underline='single')
    return c


def navigation(c,sheet,cell):
    name=sheet.replace("'","''")
    c.hyperlink=Hyperlink(ref=c.coordinate,location=f"'{name}'!{cell}")


def evidence_label(check):
    n=check['detail_count']
    if check['status'] in {'-','NA'}:
        return f'问题数量未确认；已提供明细 {n} 条'
    if check['status']=='✔':
        return '来源检查通过；已记录问题明细 0 条（不代表独立复核）'
    return f'已记录问题明细 {n} 条（不代表唯一问题总数）'


def style_new(c,header=False):
    c.font=Font(name='Arial',size=10,bold=header,color='FFFFFF' if header else '222222')
    c.alignment=Alignment(vertical='top',wrap_text=True)
    if header:c.fill=PatternFill('solid',fgColor='264F73')


def standalone_url(value):
    return isinstance(value,str) and re.fullmatch(r'https?://\S+',value.strip()) and bool(urlsplit(value.strip()).hostname)


def fit_row(ws,rownum,minimum=36):
    height=minimum
    for cell in ws[rownum]:
        if cell.value is None:continue
        width=ws.sheet_format.defaultColWidth or 13
        for dimension in ws.column_dimensions.values():
            lo=dimension.min or openpyxl.utils.column_index_from_string(dimension.index)
            hi=dimension.max or lo
            if lo<=cell.column<=hi:width=dimension.width;break
        lines=0
        for line in str(cell.value).split('\n'):
            length=sum(2 if unicodedata.east_asian_width(ch) in {'W','F'} else 1 for ch in line)
            lines+=max(1,math.ceil(length/max(1,width-2)))
        height=max(height,lines*(cell.font.sz or 10)*1.35+6)
    ws.row_dimensions[rownum].height=min(409,height)


def target_rows(main):
    rows={}
    for row in main.iter_rows():
        value=row[1].value if len(row)>1 else None
        if value is None:continue
        try:ident=check_id(value)
        except ValueError:continue
        if ident in rows:raise ValueError(f'duplicate template check {ident}')
        rows[ident]=row[0].row
    return rows


def write_details(w,run,touched):
    grouped=defaultdict(list)
    for ident,c in run['checks'].items():
        topic=run['profile']['topics'][c['topic']]
        grouped[topic['targets'][ident]].append(c)
    summary_locations={}
    evidence_cells=[]
    for name,checks in grouped.items():
        if name not in w.sheetnames:
            raise ValueError(f'missing destination worksheet; update explicit profile: {name}')
        ws=w[name]
        header_row=4 if name=='10. Robots.txt' else 3
        headers={c.column:str(c.value) for c in ws[header_row] if c.value is not None}
        if not headers:raise ValueError(f'missing destination headers in {name}')
        last_header=max(headers)
        all_fields=[]
        for c in checks:
            for d in c['details']:
                for field in d['values']:
                    if field not in all_fields:all_fields.append(field)
        mapping={}
        used=set()
        for col,h in headers.items():
            choices=[f for f in all_fields if norm(f)==norm(h)]
            if not choices:
                choices=[f for f in all_fields if canonical(f,run['profile'])==canonical(h,run['profile']) and not canonical(h,run['profile']).startswith('extra:')]
            if choices:
                mapping[col]=choices[0];used.add(choices[0])
        col=last_header+1
        for field in all_fields:
            if field not in used:
                mapping[col]=field
                headers[col]='Source: '+field
                col+=1
        provenance=['Source file','Source check','Source worksheet','Source row','Source result']
        for title in provenance:
            headers[col]=title;col+=1
        for col,title in headers.items():
            if col>last_header:
                c=put(ws,header_row,col,title,touched);style_new(c,True)
                ws.column_dimensions[c.column_letter].width=24
        occupied=[c.row for row in ws.iter_rows(min_row=header_row+1) for c in row if c.value is not None]
        start=max(occupied)+2 if occupied else header_row+1
        rownum=start
        for check in checks:
            for detail in check['details']:
                for col,field in mapping.items():
                    value=detail['values'].get(field)
                    is_url=canonical(field,run['profile']) in {'address','destination','sitemap'} or standalone_url(value)
                    c=put(ws,rownum,col,value,touched,url=is_url);c.alignment=Alignment(vertical='top',wrap_text=True)
                    if is_url and value is not None:evidence_cells.append([name,c.coordinate,str(value).strip()])
                meta=[check['source_file'],check['id'],detail['source_sheet'],detail['source_row'],check['raw_status']]
                for col,value in zip(range(max(headers)-4,max(headers)+1),meta):put(ws,rownum,col,value,touched)
                fit_row(ws,rownum)
                rownum+=1
        rownum+=2
        if ws._images:
            c=put(ws,rownum,1,'模板原有截图仅作参考，非本次审计证据',touched)
            style_new(c);ws.row_dimensions[rownum].height=36;rownum+=2
        c=put(ws,rownum,1,'检查项汇总（以下不是问题明细）',touched);style_new(c,True)
        rownum+=1
        extra_fields=list(dict.fromkeys(h for check in checks for h in check['extra_checklist']))
        summary_headers=['Check','Result','Recorded detail rows','Findings','Instruction','Source file','Source worksheet','Source row','Source result','Evidence date','Source check name']+['Source Checklist: '+h for h in extra_fields]
        for col,h in enumerate(summary_headers,1):
            c=put(ws,rownum,col,h,touched);style_new(c,True)
            if col>last_header:ws.column_dimensions[c.column_letter].width=48 if col==11 else 24
        fit_row(ws,rownum,30)
        rownum+=1
        for check in checks:
            count=check['detail_count'] if check['status'] in {'✔','✖'} else None
            values=[check['id'],check['status'],count,check['findings'],check['instruction'],check['source_file'],check['source_sheet'],check['source_row'],check['raw_status'],check['source_date'],check['source_check_name']]+[check['extra_checklist'].get(h) for h in extra_fields]
            for col,value in enumerate(values,1):style_new(put(ws,rownum,col,value,touched,url=standalone_url(value)))
            fit_row(ws,rownum,60)
            summary_locations[check['id']]=(name,f'A{rownum}')
            rownum+=1
        # Template A1 links already point back; only create one if absent.
        if not ws['A1'].hyperlink:
            touched.add((name,'A1'));navigation(ws['A1'],run['profile']['main_sheet'],'A1' if w[run['profile']['main_sheet']]['A1'].value else 'C1')
    return summary_locations,evidence_cells


def validate_output(path,original,touched,expected,evidence_cells,template_warnings,write_manifest):
    w=openpyxl.load_workbook(path,data_only=False)
    if w.sheetnames!=list(original):raise ValueError('template worksheet order/names changed')
    for name,b in original.items():
        s=w[name]
        if set(map(str,s.merged_cells.ranges))!=b['merged'] or str(s.data_validations)!=b['validation']:
            raise ValueError(f'template merges/validation changed: {name}')
        if len(s._images)!=b['images'] or len(s._charts)!=b['charts']:
            raise ValueError(f'template images/charts changed: {name}')
        for coord,signature in b['cells'].items():
            if (name,coord) not in touched and cell_signature(s[coord])!=signature:
                raise ValueError(f'unexpected template change: {name}!{coord}')
    for name,coord,value in expected:
        if (w[name][coord].value or None)!=(value or None):raise ValueError(f'output value mismatch {name}!{coord}')
    for name,coord,value,links in write_manifest:
        c=w[name][coord]
        if (c.value or None)!=(value or None) or link_signature(c)!=links:
            raise ValueError(f'written detail/navigation did not survive export: {name}!{coord}')
    for name,coord,url in evidence_cells:
        c=w[name][coord]
        if c.value!=url or not c.hyperlink or c.hyperlink.target!=url:
            raise ValueError(f'website URL hyperlink missing: {name}!{coord}')
    new_warnings=set(diagnostics(w))-set(template_warnings)
    if new_warnings:raise ValueError(f'new formula/link errors: {sorted(new_warnings)}')
    count=validate_links(w)
    w.close()
    return count


def build(manifest_path):
    run=read_run(manifest_path)
    m=run['manifest']
    output=run['folder']/f'Autocheck_onsite_audit_{m["company"]}_{m["output_date"]}.xlsx'
    sidecar=output.with_suffix('.validation.json')
    if output.exists() or sidecar.exists():raise ValueError(f'output already exists: {output}')
    w=openpyxl.load_workbook(io.BytesIO(run['template'].read_bytes()),data_only=False)
    main_name=run['profile']['main_sheet']
    if main_name not in w.sheetnames:raise ValueError(f'missing main worksheet {main_name}')
    main=w[main_name]
    original=baseline(w)
    warnings=diagnostics(w)
    # Broken inherited heading anchors are reported, not silently repaired.
    validate_links(w)
    touched=set()
    locations,evidence_cells=write_details(w,run,touched)
    rows=target_rows(main)
    expected=[]
    missing=[]
    mapped_ids=set(i for topic in run['profile']['topics'].values() for i in topic['ids'])
    status_cells=set()
    for validation in main.data_validations.dataValidation:
        if validation.type=='list' and all(x in str(validation.formula1) for x in ['✔','✖','NA','-']):
            status_cells.update(f'D{r}' for r in rows.values() if f'D{r}' in validation.sqref)
    for ident in rows:
        if ident.startswith('7.'):continue
        if ident in mapped_ids or f'D{rows[ident]}' in status_cells:
            if ident not in rows:continue
            r=rows[ident]
            check=run['checks'].get(ident)
            if check:
                findings=check['findings']+'\n[汇总] '+evidence_label(check)
                if check['status']=='-':findings+='；来源状态：'+check['raw_status']
                values=[check['status'],findings,check['instruction']]
            else:
                values=['-','未检查：本次未提供该检查项报告。','']
                missing.append(ident)
            for col,value in zip([4,5,6],values):
                c=put(main,r,col,value,touched)
                if col in {5,6}:c.alignment=copy(c.alignment);c.alignment=Alignment(horizontal=c.alignment.horizontal,vertical='top',wrap_text=True)
                expected.append([main_name,c.coordinate,value])
            if check:
                sheet,cell=locations[ident]
                c=main.cell(r,7)
                if c.data_type=='f':raise ValueError(f'navigation is a formula; explicit adapter required: {c.coordinate}')
                if c.value is None:put(main,r,7,sheet,touched)
                touched.add((main_name,c.coordinate));navigation(c,sheet,cell)
            fit_row(main,r,max(60,main.row_dimensions[r].height or 0))
    absent=set(run['checks'])-set(rows)
    if absent:raise ValueError(f'checks missing from template: {sorted(absent)}')
    w.active=w.sheetnames.index(main_name)
    fd,tmp=tempfile.mkstemp(prefix='.consolidating-',suffix='.xlsx',dir=run['folder']);os.close(fd)
    temp_path=Path(tmp)
    write_manifest=[[name,coord,w[name][coord].value,link_signature(w[name][coord])] for name,coord in sorted(touched)]
    try:
        w.save(temp_path);w.close()
        link_count=validate_output(temp_path,original,touched,expected,evidence_cells,warnings,write_manifest)
        for source in run['sources']:
            if sha(Path(source['path']).read_bytes())!=source['sha256']:raise ValueError('source changed during build')
        for p,digest in [(run['template'],run['template_sha256']),(Path(run['manifest_path']),run['manifest_sha256']),(Path(run['profile_path']),run['profile_sha256'])]:
            if sha(p.read_bytes())!=digest:raise ValueError(f'configuration changed during build: {p}')
        report={'schema_version':1,'run_id':m['run_id'],'company':m['company'],'site_hosts':m['site_hosts'],'output_date':m['output_date'],'created_utc':datetime.now(timezone.utc).isoformat(),'output_sha256':sha(temp_path.read_bytes()),'template_sha256':run['template_sha256'],'manifest_sha256':run['manifest_sha256'],'profile_sha256':run['profile_sha256'],'sources':run['sources'],'missing_checks':missing,'checks':{i:{'status':c['status'],'raw_status':c['raw_status'],'recorded_detail_rows':c['detail_count'],'source_file':c['source_file'],'summary_target':locations[i]} for i,c in run['checks'].items()},'validated_internal_links':link_count,'validated_website_links':len(evidence_cells),'template_warnings':warnings,'input_warnings':run['warnings'],'native_excel_recalculation':'not performed'}
        # Windows rename refuses an existing destination; POSIX link provides that same guarantee.
        try:
            if os.name=='nt':os.rename(temp_path,output)
            else:os.link(temp_path,output)
        except FileExistsError:raise ValueError(f'output already exists: {output}')
        try:
            with sidecar.open('x',encoding='utf-8') as f:json.dump(report,f,ensure_ascii=False,indent=2)
        except Exception:
            output.unlink();raise
    finally:
        temp_path.unlink(missing_ok=True)
    return output


def verify(path):
    path=Path(path)
    report=json.loads(path.with_suffix('.validation.json').read_text(encoding='utf-8'))
    if sha(path.read_bytes())!=report['output_sha256']:raise ValueError('output hash changed since validation')
    w=openpyxl.load_workbook(path,data_only=False)
    links=validate_links(w)
    w.close()
    return {'output':str(path),'sha256':report['output_sha256'],'checks':len(report['checks']),'internal_links':links,'website_links':report['validated_website_links'],'template_warning_count':len(report['template_warnings'])}
