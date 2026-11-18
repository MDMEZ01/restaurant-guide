import openpyxl, math
import sys
p=sys.argv[1] if len(sys.argv)>1 else '/mnt/user-data/outputs/Best Restaurants by City.xlsx'
wb=openpyxl.load_workbook(p)
def est(ws,r):
    mx=1
    for c in ws[r]:
        if c.value is None or not (c.alignment and c.alignment.wrap_text): continue
        if any(c.coordinate in m for m in ws.merged_cells.ranges): continue
        w=ws.column_dimensions[c.column_letter].width or 8.43
        sz=c.font.sz or 10
        cpl=max(1,int(w*0.95*10/sz))
        mx=max(mx,sum(max(1,math.ceil(len(s)/cpl)) for s in str(c.value).split('\n')))
    sz=max((c.font.sz or 10) for c in ws[r] if c.value is not None)
    return mx, sz
changed={}
for ws in wb:
    start=3 if ws.title!='Deep Canyon Delivery' else 4
    for r in range(start,ws.max_row+1):
        if all(c.value is None for c in ws[r]): continue
        if ws.cell(r,1).value and str(ws.cell(r,1).value).startswith('  '): continue  # section bands
        if str(ws.cell(r,1).value or '').startswith('©'): continue
        mx,_=est(ws,r)
        h=max(18, round(12.5*mx+6,1))
        if ws.row_dimensions[r].height != h:
            ws.row_dimensions[r].height=h; changed[ws.title]=changed.get(ws.title,0)+1
print(changed)
wb.save(p)
w=openpyxl.load_workbook(p)['Deep Canyon Delivery']
print([ (r,w.row_dimensions[r].height) for r in range(4,17)])
