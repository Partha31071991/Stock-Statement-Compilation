from fastapi import FastAPI
app = FastAPI(title="Stock Statement Compilation Health")

@app.get("/")
def health():
    return {"ok": True, "service": "stock-statement-compiler"}
