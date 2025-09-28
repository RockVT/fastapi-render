from fastapi import FastAPI

app = FastAPI()

@app.get("/")
def root():
    return {"message": "Hello from FastAPI on Render!"}

@app.post("/process_gsheet")
def process(data: dict):
    return {"received": data}
