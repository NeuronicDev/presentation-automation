# Dockerfile
FROM python:3.11-slim

# Install system packages first
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    curl \
    gnupg \
    libreoffice \
    poppler-utils \
    git \
    && apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Install Node.js (LTS version) and npm
RUN curl -fsSL https://deb.nodesource.com/setup_18.x | bash - && \
    apt-get update && \
    apt-get install -y nodejs && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Set work directory
WORKDIR /app

COPY requirements.txt .

# Install Python dependencies (git must be installed first!)
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

COPY package*.json ./

# Install frontend dependencies
RUN npm install

# Install global Office.js tools
RUN npm install -g yo generator-office

# Install ajv and ajv-draft-04 and save to package.json
RUN npm install ajv ajv-draft-04 --save

COPY . .

EXPOSE 8000

CMD ["python", "src/main.py"]