# 3-way merge of flat locale JSON by key, using git index stages. Prints keys changed on both sides.
import json,subprocess,sys
def stage(n,p):
    r=subprocess.run(['git','show',f':{n}:{p}'],capture_output=True,text=True)
    return json.loads(r.stdout) if r.returncode==0 else {}
for p in sys.argv[1:]:
    b,o,t=stage(1,p),stage(2,p),stage(3,p)
    out={}; both=[]
    keys=list(o.keys())
    # insert theirs-only keys after their predecessor in theirs
    tk=list(t.keys())
    for i,k in enumerate(tk):
        if k not in o and k not in keys:
            prev=next((tk[j] for j in range(i-1,-1,-1) if tk[j] in keys),None)
            keys.insert(keys.index(prev)+1 if prev else len(keys),k)
    for k in keys:
        inb,ino,inT=k in b,k in o,k in t
        bv,ov,tv=b.get(k),o.get(k),t.get(k)
        if ino and inT:
            if ov==tv or tv==bv: out[k]=ov
            elif ov==bv: out[k]=tv
            else: out[k]=tv; both.append((k,bv,ov,tv))
        elif ino:  # missing in theirs
            if inb and ov==bv: continue  # theirs deleted
            if inb: both.append((k,bv,ov,None))
            out[k]=ov
        else:  # only theirs
            if inb and tv==bv: continue  # ours deleted
            if inb: both.append((k,bv,None,tv))
            out[k]=tv
    open(p,'w').write(json.dumps(out,ensure_ascii=False,indent=2)+'\n')
    for x in both: print(p,'BOTH:',json.dumps(x,ensure_ascii=False))
