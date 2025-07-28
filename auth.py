# auth.py

import os
import bcrypt
import streamlit as st
from pymongo import MongoClient
from datetime import datetime
from dotenv import load_dotenv
import ssl
ssl._create_default_https_context = ssl._create_unverified_context

# === Load environment variables from .env ===
load_dotenv()

# === Connect to MongoDB Atlas ===
client = MongoClient(os.getenv("MONGO_URI"))
db = client["streamlit_app"]  # Name of your DB (can be anything)
users_col = db["users"]       # Collection for storing user data

# === Register new user ===
def sign_up_user(username, email, password):
    # Check if username already exists
    if users_col.find_one({"username": username}):
        return False, "Username already exists"

    # Hash the password
    hashed_pw = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt())

    # Insert user into MongoDB
    users_col.insert_one({
        "username": username,
        "email": email,
        "hashed_password": hashed_pw,
        "created_at": datetime.utcnow(),
        "role": "user"
    })
    return True, "Registration successful"

# === Authenticate login ===
def verify_user(username, password):
    user = users_col.find_one({"username": username})
    if user and bcrypt.checkpw(password.encode("utf-8"), user["hashed_password"]):
        return user
    return None

# === Log out ===
def logout():
    if "user" in st.session_state:
        del st.session_state["user"]
        st.rerun()