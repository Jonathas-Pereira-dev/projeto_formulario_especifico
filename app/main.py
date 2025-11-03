from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from app.utils import carregar_itens, salvar_resultados
from datetime import datetime

app = FastAPI()
app.mount("/static", StaticFiles(directory="app/static"), name="static")
templates = Jinja2Templates(directory="app/templates")

PLANILHA = "FT-5.82.AD.BA6XX-403 - ANEXO 1.xlsm"

@app.get("/", response_class=HTMLResponse)
async def exibir_formulario(request: Request):
    itens = carregar_itens(PLANILHA)
    return templates.TemplateResponse("formulario.html", {"request": request, "itens": itens})

@app.post("/enviar")
async def enviar_formulario(request: Request):
    form = await request.form()
    dados = []
    
    for key in form:
        if key.startswith("status_"):
            item_id = key.split("_")[1]
            status = form[key]
            justificativa = form.get(f"just_{item_id}", "")
            equipamento = form.get(f"equipamento_{item_id}", "")
            aba = form.get(f"aba_{item_id}", "")
            
            dados.append({
                "Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Aba": aba,
                "Equipamento": equipamento,
                "Status": status,
                "Justificativa": justificativa
            })
    
    # Gera um nome de arquivo com a data atual
    nome_arquivo = f"resultados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    salvar_resultados(nome_arquivo, dados)
    return RedirectResponse(url="/", status_code=303)
