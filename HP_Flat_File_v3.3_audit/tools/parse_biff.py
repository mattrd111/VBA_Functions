import struct, os, re, sys

def records(data):
    i=0; n=len(data)
    while i<n:
        # record id: 1-2 bytes
        b=data[i]; i+=1
        rid=b&0x7F
        if b&0x80:
            b2=data[i]; i+=1
            rid|=(b2&0x7F)<<7
        # size: up to 4 bytes 7-bit
        size=0; shift=0
        for k in range(4):
            b=data[i]; i+=1
            size|=(b&0x7F)<<shift; shift+=7
            if not b&0x80: break
        yield rid, data[i:i+size]
        i+=size

def wstr(buf, off):
    cch=struct.unpack_from('<I',buf,off)[0]
    if cch==0xFFFFFFFF: return None, off+4
    s=buf[off+4:off+4+2*cch].decode('utf-16le',errors='replace')
    return s, off+4+2*cch

if __name__=="__main__":
    wb=open('xlsb/xl/workbook.bin','rb').read()
    sheets=[]
    for rid,body in records(wb):
        if rid==0x9C: # BrtBundleSh
            hs,itab=struct.unpack_from('<II',body,0)
            rel,off=wstr(body,8)
            name,off=wstr(body,off)
            sheets.append((itab,rel,name,hs))
    rels=open('xlsb/xl/_rels/workbook.bin.rels',encoding='utf-8').read()
    relmap=dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"',rels))
    relmap.update({k:v for v,k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"',rels)})
    print("itab | rId | target | state | name | codeName")
    for itab,rel,name,hs in sheets:
        tgt=relmap.get(rel,'?')
        code=None
        p=os.path.join('xlsb/xl',tgt)
        if os.path.exists(p):
            d=open(p,'rb').read(20000)
            for rid,body in records(d):
                if rid==0x93: # BrtWsProp
                    try:
                        code,_=wstr(body,19)
                    except Exception as e:
                        code='ERR'
                    break
        print(itab,rel,tgt,hs,repr(name),repr(code))
