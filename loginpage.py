import streamlit as st
import os
import datetime
import cv2
import face_recognition

if "users" not in st.session_state:
    st.session_state.users = {}

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "username" not in st.session_state:
    st.session_state.username = ""



if st.session_state.logged_in:
    menu = st.sidebar.selectbox("Menu", ["Take Photo", "Logout","Mark Attendence"])
else:
    menu = st.sidebar.selectbox("Menu", ["Login", "Signup"])


st.title("REAL-TIME FACE ATTENDANCE SYSTEM")



if menu == "Signup":
    st.subheader("Create New Account")

    new_user = st.text_input("Username")
    new_password = st.text_input("Password", type="password")

    if st.button("Signup"):
        if new_user in st.session_state.users:
            st.warning("User already exists!")
        else:
            st.session_state.users[new_user] = new_password
            st.success("Account created successfully!")

elif menu == "Login":
    st.subheader("Login to your account")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username in st.session_state.users and st.session_state.users[username] == password:
            st.session_state.logged_in = True
            st.session_state.username = username

            st.success(f"Welcome {username} 🎉")
            st.rerun()   
        else:
            st.error("Invalid username or password")



elif menu == "Take Photo":
    if st.session_state.logged_in:
        st.subheader("📸 Take Photo for Attendance")

        img = st.camera_input("Capture your face")

        if img is not None:
            st.image(img, caption="Captured Image")

            folder = "D:/photo_save"
            os.makedirs(folder, exist_ok=True)

            time_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{st.session_state.username}_{time_str}.jpg"
            file_path = os.path.join(folder, filename)

            with open(file_path, "wb") as f:
                f.write(img.getbuffer())

            st.success(f"Image saved at {file_path}")
    else:
        st.warning("Please login first!")
# elif menu=="Mark Attendence":
#     if st.session_state.logged_in:
        

elif menu == "Logout":
    st.session_state.logged_in = False
    st.session_state.username = ""

    st.success("Logged out successfully")
    st.rerun()   