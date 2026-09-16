from __future__ import annotations
import os, tempfile, uuid
from flask import Flask, render_template, request, redirect, url_for, abort
from analyzer import analyze, LABELS

app=Flask(__name__); app.config["MAX_CONTENT_LENGTH"]=50*1024*1024
RESULTS={}
@app.get("/")
def index(): return render_template("index.html")
@app.post("/analyze")
def upload():
    file=request.files.get("file")
    if not file or not file.filename.lower().endswith(".xlsx"): return render_template("index.html",error="Seleccione un archivo .xlsx válido."),400
    fd,path=tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
    try:
        file.save(path); result=analyze(path)
    except Exception as e:
        return render_template("index.html",error=f"No fue posible leer el libro: {e}"),400
    finally: os.unlink(path)
    token=str(uuid.uuid4()); RESULTS[token]=result
    return redirect(url_for("groups",token=token))
def result_or_404(token):
    if token not in RESULTS: abort(404)
    return RESULTS[token]
@app.get("/result/<token>")
def groups(token):
    r=result_or_404(token); grouped=[]
    for kind,count in r["groups"].items():
        grouped.append({"kind":kind,"label":LABELS[kind],"count":count,"evidence_count":r["evidence_counts"][kind]})
    return render_template("groups.html",result=r,groups=grouped,token=token)
@app.get("/result/<token>/group/<kind>")
def incident_list(token,kind):
    r=result_or_404(token); return render_template("incidents.html",token=token,kind=kind,label=LABELS.get(kind,kind),incidents=[x for x in r["incidents"] if x["kind"]==kind])
@app.get("/result/<token>/incident/<int:incident_id>")
def detail(token,incident_id):
    r=result_or_404(token)
    if incident_id>=len(r["incidents"]): abort(404)
    inc=r["incidents"][incident_id]
    from analyzer import rows_for
    return render_template("detail.html",token=token,inc=inc,constructora=rows_for(inc.get("evidence",{}).get("constructora",{"dates":[],"receipts":[],"values":[]})),fiduciaria=rows_for(inc.get("evidence",{}).get("fiduciaria",{"dates":[],"receipts":[],"values":[]})))
if __name__=="__main__": app.run(host="127.0.0.1",port=5000,debug=False)
