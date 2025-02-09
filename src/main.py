from typing import Annotated
from fastapi import FastAPI, Request, Form, UploadFile
from fastapi.templating import Jinja2Templates
from .attendance import cal_attendance
from fastapi.responses import FileResponse

templates = Jinja2Templates(directory="src/templates")


app = FastAPI()


@app.get("/")
def homepage_get(request: Request):
    name = "Hello World"
    return templates.TemplateResponse(
        request=request, name="page.html", context={"name": name}
    )


@app.post("/")
def homepage_post(name: Annotated[str, Form()], request: Request):
    print(name)
    return templates.TemplateResponse(
        request=request, name="page.html", context={"name": name}
    )


@app.post("/uploadfile/")
async def create_upload_file(file: UploadFile):
    try:
        cal_attendance(file)
        return FileResponse("./src/temp/out.xlsx")
    except:
        return {"error": "error"}
