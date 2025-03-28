FROM python:3.10-slim

WORKDIR /app

RUN pip install --no-cache-dir python-pptx

COPY executor_script.py .

CMD ["python", "executor_script.py"]


# FROM python:3.10-slim

# WORKDIR /app

# RUN pip install --no-cache-dir python-pptx

# # Explicitly set shell to ensure environment is loaded
# SHELL ["/bin/bash", "-c"]

# COPY executor_script.py .

# # Ensure the script is executable
# RUN chmod +x executor_script.py

# # Use explicit entrypoint to ensure environment is loaded
# ENTRYPOINT ["/bin/bash", "-c"]
# CMD ["python executor_script.py"]