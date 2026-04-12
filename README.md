# Notas de Memoria Financiera - TickerWEB

Aplicacion web que genera notas de memoria financiera profesionales a partir de datos reales de Yahoo Finance, con redaccion automatica mediante inteligencia artificial (Claude de Anthropic).

## Funcionalidades

- Consulta de datos financieros reales por ticker (Yahoo Finance)
- Tabla profesional con ingresos, margenes, EBITDA, beneficio neto y porcentajes
- Generacion automatica de nota explicativa en espanol con lenguaje contable profesional
- Descarga de la nota completa en formato Word (.docx)

## Requisitos

- Python 3.9 o superior
- Una API key de Anthropic (https://console.anthropic.com/)

## Instalacion

1. Clona el repositorio:

```bash
git clone https://github.com/victorargenta/notas-memoria-financiera-tickerweb.git
cd notas-memoria-financiera-tickerweb
```

2. Crea y activa un entorno virtual:

```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate     # Windows
```

3. Instala las dependencias:

```bash
pip install -r requirements.txt
```

4. Configura las variables de entorno:

```bash
cp .env.example .env
```

Edita el archivo `.env` y anade tu API key de Anthropic:

```
ANTHROPIC_API_KEY=sk-ant-tu-clave-aqui
```

5. (Opcional) Anade tu logo corporativo como `static/logo.png`.

## Ejecucion

```bash
python app.py
```

La aplicacion estara disponible en: **http://localhost:5000**

## Uso

1. Introduce el ticker de la empresa (por ejemplo: `AAPL`, `MSFT`, `TSLA`, `BBVA.MC`)
2. Pulsa "Analizar"
3. Revisa la tabla de datos financieros y la nota generada
4. Descarga la nota en formato Word con el boton "Descargar Word"

## Estructura del proyecto

```
.
├── app.py                 # Aplicacion Flask principal
├── requirements.txt       # Dependencias Python
├── .env.example           # Plantilla de variables de entorno
├── .gitignore             # Archivos excluidos de git
├── static/
│   └── logo.png           # Logo corporativo (opcional)
├── templates/
│   └── index.html         # Plantilla HTML principal
└── README.md              # Este archivo
```

## Tecnologias

- **Flask** - Framework web
- **yfinance** - Datos financieros de Yahoo Finance
- **Anthropic API** - Generacion de texto con Claude
- **python-docx** - Generacion de documentos Word
- **python-dotenv** - Gestion de variables de entorno
