from pathlib import Path

from fastapi import FastAPI, Request, File, UploadFile, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.cors import CORSMiddleware
from mangum import Mangum

from csvtodocx import convertCSV

app = FastAPI()
handler = Mangum(app)

templates = Jinja2Templates(directory=Path(
    Path(__file__).parent.parent.absolute() / 'src/templates'))

origins = ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount(
    "/static",
    StaticFiles(directory=Path(__file__).parent.parent.absolute() /
                'src/static'),  # Path(__file__).parent.parent.absolute()
    name="static",
)

# -------------------------------------HTML Templates-------------------------------------


@app.get('/')
def login(req: Request):
    return templates.TemplateResponse('index.html', {'request': req})


# --------------------------------------API Requests--------------------------------------
@app.post("/convert_to_docx")
async def upload_file(request: Request, csv: UploadFile = File(...), model: UploadFile = File(...)):
    zip_bytes = await convertCSV(csv, model)
    return Response(content=zip_bytes, media_type="application/zip")


# To run the app for development: python -m uvicorn server:app --host 0.0.0.0 --reload
# To run the app on AWS: python3 -m uvicorn server:app