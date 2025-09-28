from fastapi import FastAPI
import mycode   # import file mycode.py

app = FastAPI()

@app.get("/")
def root():
    return {"message": "FastAPI is running"}

@app.get("/run_mycode")
def run_mycode():
    result = mycode.run_mycode()   # gọi hàm từ file mycode.py
    return {"result": result}
