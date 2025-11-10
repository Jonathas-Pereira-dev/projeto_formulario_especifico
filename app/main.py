import os
from datetime import datetime, timedelta
from typing import Optional

from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

# Carrega variáveis de ambiente
from dotenv import load_dotenv
load_dotenv(dotenv_path=".env")

from jose import JWTError, jwt
from passlib.context import CryptContext

# SQLAlchemy
from sqlalchemy import Column, Integer, String, Boolean, create_engine, or_
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# Funções utilitárias
from app.utils import carregar_itens, salvar_resultados, carregar_abas

# CONFIGURAÇÕES GERAIS

SECRET_KEY = os.getenv("SECRET_KEY", "mude_para_producao_gerar_com_openssl")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24

# Banco de dados automático (Railway ou local)
DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL or "railway.internal" in DATABASE_URL:
    print(" Usando banco local SQLite (modo desenvolvimento)")
    DATABASE_URL = "sqlite:///./users.db"
else:
    print(f" Usando banco remoto: {DATABASE_URL}")

COOKIE_NAME = "access_token"
COOKIE_SECURE = os.getenv("COOKIE_SECURE", "false").lower() == "true"

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# BANCO DE DADOS

connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, connect_args=connect_args)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
Base = declarative_base()


class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True, nullable=False)
    email = Column(String, unique=True, index=True, nullable=True)
    hashed_password = Column(String, nullable=False)
    is_active = Column(Boolean, default=True)


Base.metadata.create_all(bind=engine)

# APLICAÇÃO FASTAPI

app = FastAPI(title="Radix - Inspeção (com Auth)")
app.mount("/static", StaticFiles(directory="app/static"), name="static")
templates = Jinja2Templates(directory="app/templates")

PLANILHA = "app/static/planilhas/FT-5.82.AD.BA6XX-403 - ANEXO 1.xlsm"

# AUTENTICAÇÃO

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def verify_password(plain_password, hashed_password):
    try:
        return pwd_context.verify(plain_password, hashed_password)
    except Exception:
        return False


def get_password_hash(password):
    # bcrypt aceita no máximo 72 bytes — truncamos se exceder
    if len(password) > 72:
        password = password[:72]
    return pwd_context.hash(password)


def create_access_token(data: dict, expires_delta: Optional[timedelta] = None):
    to_encode = data.copy()
    expire = datetime.utcnow() + (expires_delta or timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES))
    to_encode.update({"exp": expire})
    return jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)


def decode_token_get_username(token: str) -> Optional[str]:
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        return payload.get("sub")
    except JWTError:
        return None


# ROTAS DE LOGIN / REGISTRO

@app.get("/login", response_class=HTMLResponse)
async def get_login(request: Request):
    return templates.TemplateResponse("login.html", {"request": request})

@app.post("/login")
async def post_login(request: Request, username: str = Form(...), password: str = Form(...)):
    db = next(get_db())
    user = db.query(User).filter(User.username == username).first()
    if not user or not verify_password(password, user.hashed_password):
        return templates.TemplateResponse(
            "login.html",
            {"request": request, "error": "Usuário ou senha inválidos"},
        )
    token = create_access_token({"sub": user.username})
    response = RedirectResponse(url="/", status_code=303)
    response.set_cookie(
        COOKIE_NAME, token, httponly=True, secure=COOKIE_SECURE, samesite="lax"
    )
    return response

@app.get("/register", response_class=HTMLResponse)
async def get_register(request: Request):
    return templates.TemplateResponse("register.html", {"request": request})

@app.post("/register")
async def post_register(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    email: str = Form(None),
):
    db = next(get_db())
    existing_user = db.query(User).filter(or_(User.username == username, User.email == email)).first()
    if existing_user:
        return templates.TemplateResponse(
            "register.html",
            {"request": request, "error": "Usuário ou e-mail já cadastrado"},
        )

    if len(password) > 72:
        return templates.TemplateResponse(
            "register.html",
            {"request": request, "error": "A senha não pode ultrapassar 72 caracteres."},
        )

    user = User(username=username, email=email, hashed_password=get_password_hash(password))
    db.add(user)
    db.commit()
    return RedirectResponse(url="/login", status_code=303)

@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(COOKIE_NAME)
    return response


# OBTÉM USUÁRIO LOGADO

def get_current_user_from_request(request: Request):
    token = request.cookies.get(COOKIE_NAME)
    if not token:
        return None
    username = decode_token_get_username(token)
    if not username:
        return None
    db = next(get_db())
    return db.query(User).filter(User.username == username).first()

# ROTAS PRINCIPAIS (PROTEGIDAS)

@app.get("/", response_class=HTMLResponse)
async def pagina_inicial(request: Request):
    user = get_current_user_from_request(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    abas = carregar_abas(PLANILHA)
    return templates.TemplateResponse(
        "index.html", {"request": request, "abas": abas, "user": {"username": user.username}}
    )

@app.get("/formulario/{aba_id}", response_class=HTMLResponse)
async def exibir_formulario(request: Request, aba_id: str):
    user = get_current_user_from_request(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    result = carregar_itens(PLANILHA, aba_id)
    itens = result.get("items", []) if isinstance(result, dict) else result
    headers = result.get("headers", []) if isinstance(result, dict) else []
    return templates.TemplateResponse(
        "formulario.html",
        {
            "request": request,
            "itens": itens,
            "headers": headers,
            "aba_id": aba_id,
            "user": {"username": user.username},
        },
    )

@app.get("/inspecao-visual", response_class=HTMLResponse)
async def ir_para_inspecao_visual(request: Request):
    return RedirectResponse(url="/formulario/2", status_code=303)

@app.post("/enviar")
async def enviar_formulario(request: Request):
    user = get_current_user_from_request(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    form = await request.form()
    dados = []

    for key in form:
        if key.startswith("status_"):
            item_id = key.split("_")[1]
            status = form[key]
            justificativa = form.get(f"just_{item_id}", "")
            equipamento = form.get(f"equipamento_{item_id}", "")
            aba = form.get(f"aba_{item_id}", "")
            dados.append(
                {
                    "Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Usuario": user.username,
                    "Aba": aba,
                    "Equipamento": equipamento,
                    "Status": status,
                    "Justificativa": justificativa,
                }
            )

    nome_arquivo = f"resultados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    salvar_resultados(nome_arquivo, dados)
    return RedirectResponse(url="/", status_code=303)