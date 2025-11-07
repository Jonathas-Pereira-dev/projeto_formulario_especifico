import os
from datetime import datetime, timedelta
from typing import Optional

from fastapi import FastAPI, Request, Form, HTTPException
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

# Carrega variáveis de ambiente apenas em desenvolvimento
from dotenv import load_dotenv

# Só carrega .env se não estiver no Railway (produção)
if not os.getenv("RAILWAY_ENVIRONMENT"):
    load_dotenv()

from jose import JWTError, jwt
from passlib.context import CryptContext

# SQLAlchemy
from sqlalchemy import Column, Integer, String, Boolean, create_engine, or_
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

# CONFIGURAÇÕES GERAIS

SECRET_KEY = os.getenv("SECRET_KEY", "chave_temporaria_para_desenvolvimento")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24

# Banco de dados - Correção para Railway
DATABASE_URL = os.getenv("DATABASE_URL")

# CORREÇÃO: Detecta PostgreSQL corretamente e trata SQLite
if DATABASE_URL and "postgresql" in DATABASE_URL:
    print("🔹 Usando banco PostgreSQL do Railway")
else:
    print("🔸 Usando banco local SQLite (modo desenvolvimento)")
    DATABASE_URL = "sqlite:///./users.db"

COOKIE_NAME = "access_token"
COOKIE_SECURE = os.getenv("COOKIE_SECURE", "true").lower() == "true"

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

# BANCO DE DADOS

try:
    if DATABASE_URL.startswith("sqlite"):
        connect_args = {"check_same_thread": False}
        engine = create_engine(DATABASE_URL, connect_args=connect_args)
    else:
        engine = create_engine(DATABASE_URL)
    
    SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)
    Base = declarative_base()
    
    class User(Base):
        __tablename__ = "users"
        id = Column(Integer, primary_key=True, index=True)
        username = Column(String, unique=True, index=True, nullable=False)
        email = Column(String, unique=True, index=True, nullable=True)
        hashed_password = Column(String, nullable=False)
        is_active = Column(Boolean, default=True)
    
    # Cria tabelas
    Base.metadata.create_all(bind=engine)
    print("✅ Banco de dados configurado com sucesso")
    
except Exception as e:
    print(f"❌ Erro crítico no banco de dados: {e}")
    # Fallback para evitar que a aplicação quebre completamente
    engine = None
    SessionLocal = None
    Base = None
    User = None

# APLICAÇÃO FASTAPI

app = FastAPI(title="Radix - Inspeção (com Auth)")

# Configurações de arquivos estáticos com fallback
try:
    app.mount("/static", StaticFiles(directory="app/static"), name="static")
    templates = Jinja2Templates(directory="app/templates")
    print("✅ Templates e static files configurados")
except Exception as e:
    print(f"⚠️  Aviso em arquivos estáticos: {e}")
    templates = None

PLANILHA = "FT-5.82.AD.BA6XX-403 - ANEXO 1.xlsm"

# AUTENTICAÇÃO

def get_db():
    if SessionLocal is None:
        yield None
        return
        
    db = SessionLocal()
    try:
        yield db
    except Exception as e:
        print(f"❌ Erro de conexão com banco: {e}")
        yield None
    finally:
        try:
            db.close()
        except:
            pass

def verify_password(plain_password, hashed_password):
    try:
        return pwd_context.verify(plain_password, hashed_password)
    except Exception:
        return False

def get_password_hash(password):
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

# ROTA DE HEALTH CHECK (IMPORTANTE PARA RAILWAY)
@app.get("/")
async def health_check():
    return {"status": "online", "message": "Aplicação rodando com sucesso"}

# ROTAS DE LOGIN / REGISTRO

@app.get("/login", response_class=HTMLResponse)
async def get_login(request: Request):
    if templates is None:
        return HTMLResponse("<h1>Sistema em manutenção</h1>")
    return templates.TemplateResponse("login.html", {"request": request})

@app.post("/login")
async def post_login(request: Request, username: str = Form(...), password: str = Form(...)):
    if templates is None:
        return HTMLResponse("<h1>Sistema em manutenção</h1>")
        
    db = next(get_db())
    if db is None:
        return templates.TemplateResponse(
            "login.html",
            {"request": request, "error": "Sistema temporariamente indisponível"},
        )
    
    user = db.query(User).filter(User.username == username).first()
    if not user or not verify_password(password, user.hashed_password):
        return templates.TemplateResponse(
            "login.html",
            {"request": request, "error": "Usuário ou senha inválidos"},
        )
    
    token = create_access_token({"sub": user.username})
    response = RedirectResponse(url="/home", status_code=303)
    response.set_cookie(
        COOKIE_NAME, token, httponly=True, secure=COOKIE_SECURE, samesite="lax"
    )
    return response

@app.get("/register", response_class=HTMLResponse)
async def get_register(request: Request):
    if templates is None:
        return HTMLResponse("<h1>Sistema em manutenção</h1>")
    return templates.TemplateResponse("register.html", {"request": request})

@app.post("/register")
async def post_register(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    email: str = Form(None),
):
    if templates is None:
        return HTMLResponse("<h1>Sistema em manutenção</h1>")
        
    db = next(get_db())
    if db is None:
        return templates.TemplateResponse(
            "register.html",
            {"request": request, "error": "Sistema temporariamente indisponível"},
        )
    
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

    try:
        user = User(username=username, email=email, hashed_password=get_password_hash(password))
        db.add(user)
        db.commit()
        return RedirectResponse(url="/login", status_code=303)
    except Exception as e:
        db.rollback()
        return templates.TemplateResponse(
            "register.html",
            {"request": request, "error": f"Erro ao criar usuário: {str(e)}"},
        )

@app.get("/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=303)
    response.delete_cookie(COOKIE_NAME)
    return response

# ROTAS DE DEBUG
@app.get("/debug/db")
async def debug_database():
    try:
        db = next(get_db())
        if db is None:
            return {"status": "error", "error": "Não foi possível conectar ao banco"}
        user_count = db.query(User).count()
        return {
            "status": "success", 
            "database_url": DATABASE_URL,
            "user_count": user_count,
            "tables_created": True
        }
    except Exception as e:
        return {"status": "error", "error": str(e)}

@app.get("/debug/env")
async def debug_env():
    return {
        "secret_key_configured": bool(SECRET_KEY and SECRET_KEY != "chave_temporaria_para_desenvolvimento"),
        "cookie_secure": COOKIE_SECURE,
        "database_url": DATABASE_URL,
        "railway_environment": bool(os.getenv("RAILWAY_ENVIRONMENT"))
    }

# OBTÉM USUÁRIO LOGADO
def get_current_user_from_request(request: Request):
    token = request.cookies.get(COOKIE_NAME)
    if not token:
        return None
    username = decode_token_get_username(token)
    if not username:
        return None
    db = next(get_db())
    if db is None:
        return None
    return db.query(User).filter(User.username == username).first()

# ROTAS PRINCIPAIS (PROTEGIDAS)
@app.get("/home", response_class=HTMLResponse)
async def pagina_inicial(request: Request):
    user = get_current_user_from_request(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    
    if templates is None:
        return HTMLResponse("<h1>Bem-vindo {user.username}</h1><p>Sistema carregado</p>")
    
    try:
        from app.utils import carregar_abas
        abas = carregar_abas(PLANILHA)
        return templates.TemplateResponse(
            "index.html", {"request": request, "abas": abas, "user": {"username": user.username}}
        )
    except Exception as e:
        return templates.TemplateResponse(
            "index.html", {"request": request, "abas": [], "user": {"username": user.username}, "error": str(e)}
        )

@app.get("/formulario/{aba_id}", response_class=HTMLResponse)
async def exibir_formulario(request: Request, aba_id: str):
    user = get_current_user_from_request(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)
    
    if templates is None:
        return HTMLResponse("<h1>Formulário</h1>")
    
    try:
        from app.utils import carregar_itens
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
    except Exception as e:
        return templates.TemplateResponse(
            "error.html", {"request": request, "error": str(e)}
        )

@app.get("/inspecao-visual", response_class=HTMLResponse)
async def ir_para_inspecao_visual(request: Request):
    return RedirectResponse(url="/formulario/2", status_code=303)

@app.post("/enviar")
async def enviar_formulario(request: Request):
    user = get_current_user_from_request(request)
    if not user:
        return RedirectResponse(url="/login", status_code=303)

    try:
        from app.utils import salvar_resultados
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
        return RedirectResponse(url="/home", status_code=303)
    except Exception as e:
        return RedirectResponse(url="/home?error=" + str(e), status_code=303)

# IMPORTANTE: Adicione isso para o Railway
if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)