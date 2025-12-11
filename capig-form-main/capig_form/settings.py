from pathlib import Path
import environ

# Base dir
BASE_DIR = Path(__file__).resolve().parent.parent
ENV_PATH = Path(__file__).resolve().parent / '.env'

# Inicializa entorno
env = environ.Env()

# Carga archivo .env desde la ruta explicita
environ.Env.read_env(str(ENV_PATH))

# Claves desde .env
SECRET_KEY = env.str('SECRET_KEY')
SHEET_PATH = env.str('SHEET_PATH')
SERVICE = env.str('SERVICE')
SECURITY_CODE = env.str('SECURITY_CODE', default='000000')

# Manejo robusto de entorno
ENVIRONMENT = env.str('ENVIRONMENT', default='dev').lower()
DEBUG = ENVIRONMENT != 'prod'

ALLOWED_HOSTS = ['*'] if DEBUG else ['localhost', '127.0.0.1', 'capig-form.onrender.com']
CSRF_TRUSTED_ORIGINS = ['https://capig-form.onrender.com'] if not DEBUG else []

# Aplicaciones instaladas
INSTALLED_APPS = [
    'django.contrib.admin',
    'django.contrib.auth',
    'django.contrib.contenttypes',
    'django.contrib.sessions',
    'django.contrib.messages',
    'django.contrib.staticfiles',
    'corsheaders',
    'forms',
    'django_extensions',
]

# Middleware
MIDDLEWARE = [
    'django.middleware.security.SecurityMiddleware',
    'django.contrib.sessions.middleware.SessionMiddleware',
    'corsheaders.middleware.CorsMiddleware',
    'django.middleware.common.CommonMiddleware',
    'django.middleware.csrf.CsrfViewMiddleware',
    'django.contrib.auth.middleware.AuthenticationMiddleware',
    'django.contrib.messages.middleware.MessageMiddleware',
    'django.middleware.clickjacking.XFrameOptionsMiddleware',
]

# Rutas
ROOT_URLCONF = 'capig_form.urls'

# Templates
TEMPLATES = [
    {
        'BACKEND': 'django.template.backends.django.DjangoTemplates',
        'DIRS': [],
        'APP_DIRS': True,
        'OPTIONS': {
            'context_processors': [
                'django.template.context_processors.debug',
                'django.template.context_processors.request',
                'django.contrib.auth.context_processors.auth',
                'django.contrib.messages.context_processors.messages',
            ],
        },
    },
]

# WSGI
WSGI_APPLICATION = 'capig_form.wsgi.application'

# Base de datos
DATABASES = {
    'default': {
        'ENGINE': 'django.db.backends.sqlite3',
        'NAME': BASE_DIR / 'db.sqlite3',
    }
}

# Seguridad y CORS
CORS_ALLOW_ALL_ORIGINS = True
X_FRAME_OPTIONS = 'ALLOWALL'

# ValidaciÃ³n de contraseÃ±as
AUTH_PASSWORD_VALIDATORS = [
    {'NAME': 'django.contrib.auth.password_validation.UserAttributeSimilarityValidator'},
    {'NAME': 'django.contrib.auth.password_validation.MinimumLengthValidator'},
    {'NAME': 'django.contrib.auth.password_validation.CommonPasswordValidator'},
    {'NAME': 'django.contrib.auth.password_validation.NumericPasswordValidator'},
]

# InternacionalizaciÃ³n
LANGUAGE_CODE = 'en-us'
TIME_ZONE = 'UTC'
USE_I18N = True
USE_TZ = True

# Archivos estÃ¡ticos
STATIC_URL = 'static/'

# Tipo de clave primaria por defecto
DEFAULT_AUTO_FIELD = 'django.db.models.BigAutoField'
