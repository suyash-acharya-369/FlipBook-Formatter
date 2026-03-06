import streamlit as st
import streamlit.components.v1 as components
import os
import subprocess
import time
import socket

# Set page config for a premium look
st.set_page_config(
    page_title="Flip-Book Formatter | AISECT Learn",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Hide Streamlit header, footer and menu for a pure HTML look
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            .block-container {
                padding-top: 0rem;
                padding-bottom: 0rem;
                padding-left: 0rem;
                padding-right: 0rem;
            }
            iframe {
                border: none;
                width: 100%;
                height: 100vh;
                overflow: hidden;
            }
            .stApp {
                overflow: hidden;
            }
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

def is_port_in_use(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) == 0

def main():
    # Ensure Flask backend is running on port 5055
    if not is_port_in_use(5055):
        # We start it as a background process
        # On Windows we use CREATE_NEW_PROCESS_GROUP to keep it alive
        try:
            subprocess.Popen(["python", "app.py"], 
                            cwd=os.path.dirname(os.path.abspath(__file__)),
                            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP if os.name == 'nt' else 0)
            time.sleep(3) # Give it 3 seconds to start
        except Exception as e:
            st.error(f"Failed to start backend: {e}")
            return

    # Render the original HTML interface via Iframe
    # Fixed height to 100vh and scrolling disabled for a seamless "App" feel
    components.iframe("http://localhost:5055", height=0, scrolling=False)
    
    # Use custom CSS to force the iframe to fill the height since Streamlit's height param 
    # can be finicky with responsive viewports
    st.markdown("""
        <style>
            iframe {
                position: fixed;
                top: 0;
                left: 0;
                width: 100vw;
                height: 100vh;
            }
        </style>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
