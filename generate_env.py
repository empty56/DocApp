import os
import secrets

env_path = ".env"

if not os.path.exists(env_path):
    secret_key = secrets.token_urlsafe(50)

    with open(env_path, "w", encoding="utf-8") as f:
        f.write(f"SECRET_KEY={secret_key}\n")
        f.write("DEBUG=True\n")
        f.write("ALLOWED_HOSTS=127.0.0.1,localhost\n")

    print("✅ .env file created.")
else:
    print("ℹ️ .env file already exists. Skipping.")