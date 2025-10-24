# PowerPoint Add-in Django Backend

This is the Django backend for the PowerPoint add-in project.

## Setup Instructions

### 1. Environment Setup

1. Copy the environment file:
   ```bash
   cp env.example .env
   ```

2. Update the `.env` file with your database credentials and secret key.

### 2. Database Setup

Start PostgreSQL using Docker Compose:
```bash
docker-compose up -d
```

### 3. Python Dependencies

Install the required packages:
```bash
pip install -r requirements.txt
```

### 4. Django Setup

1. Run migrations:
   ```bash
   python manage.py migrate
   ```

2. Create a superuser:
   ```bash
   python manage.py createsuperuser
   ```

3. Start the development server:
   ```bash
   python manage.py runserver
   ```

## API Endpoints

- **Health Check**: `GET /api/health/` - Returns `{"status": "ok"}`

## Features

- Django 5.0 with Django REST Framework
- PostgreSQL database with Docker Compose
- JWT Authentication
- CORS headers configured for Office.js add-in
- Media file handling for questions and answers
- Health check endpoint for monitoring

## Environment Variables

- `DB_NAME`: Database name
- `DB_USER`: Database user
- `DB_PASSWORD`: Database password
- `DB_HOST`: Database host
- `DB_PORT`: Database port
- `SECRET_KEY`: Django secret key
- `DEBUG`: Debug mode (True/False)
- `ALLOWED_HOSTS`: Comma-separated list of allowed hosts
