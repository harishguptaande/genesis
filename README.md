# Genesis - Full Stack Python Application

A modern full-stack application built with FastAPI, SQLAlchemy, and Pydantic.

## Project Structure

```
genesis/
├── app/
│   ├── __init__.py
│   ├── main.py              # FastAPI app entry point
│   ├── config.py            # Configuration management
│   ├── models/              # SQLAlchemy models
│   │   └── __init__.py
│   ├── schemas/             # Pydantic schemas
│   │   └── __init__.py
│   ├── routes/              # API routes
│   │   └── __init__.py
│   └── database.py          # Database setup
├── tests/
│   ├── __init__.py
│   ├── test_main.py         # Main app tests
│   └── conftest.py          # Pytest fixtures
├── .gitignore
├── requirements.txt
├── .env.example
├── README.md
└── pytest.ini
```

## Setup

### 1. Install Dependencies

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure Environment

Copy `.env.example` to `.env` and update values:

```bash
cp .env.example .env
```

### 3. Run Development Server

```bash
uvicorn app.main:app --reload
```

Server runs at `http://localhost:8000`

API docs: `http://localhost:8000/docs` (Swagger UI)

### 4. Run Tests

```bash
pytest
```

## Development

- **Add models**: Create in `app/models/`
- **Add schemas**: Create in `app/schemas/`
- **Add routes**: Create in `app/routes/`
- **Add tests**: Create in `tests/`

## License

MIT
