from fastapi import FastAPI
app=FastAPI(title="Stock Statement Compilation API")
@app.get("/")
def root(): return {"ok":True,"service":"stock-statement-compiler","note":"Use /api/health and /api/analyze"}
