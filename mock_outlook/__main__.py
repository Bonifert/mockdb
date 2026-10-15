import uvicorn

uvicorn.run("mock_outlook.app:app", host="127.0.0.1", port=8000, reload=True)
