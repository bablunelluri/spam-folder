import streamlit as st
import pickle
from sklearn.feature_extraction.text import CountVectorizer
from win32com.client import Dispatch
import pythoncom

def speak(text):
    speak_engine = Dispatch(("SAPI.SpVoice"))
    speak_engine.Speak(text)

def display_feedback_form():
    # Replace 'YOUR_GOOGLE_FORM_LINK' with the actual link to your Google Form
    feedback_form_link = 'https://docs.google.com/forms/d/e/1FAIpQLSc0ZE1R6ZO-6minilOvrRH6jXLRrhhheS7X4LKbt6f0ZpDEyg/viewform'
    
    # Display the Google Form using iframe
    st.markdown(f'<iframe src="{feedback_form_link}" width="800" height="600"></iframe>', unsafe_allow_html=True)

def about_section():
    st.subheader("About Email Spam Classification App")
    st.write("Welcome to the Email Spam Classification App!")
    st.write("This application uses machine learning to classify emails as spam or not spam.")
    st.write("Feel free to explore the classification feature and provide feedback to help us improve.")
    st.write("Your input is valuable!")

def footer():
    st.write("© 2024 by Tejkumar Nelluri. All rights reserved.")

def main():
    # Initialize the COM library
    pythoncom.CoInitialize()

    try:
        st.title("Email Spam Classification Application")

        # Display "Classification" and "About" buttons on the left sidebar
        st.sidebar.subheader("Navigation")
        classification_clicked = st.sidebar.button("Classification")
        about_clicked = st.sidebar.button("About")

        if classification_clicked or not about_clicked:
            # Default view: "Classification"
            st.subheader("Classification")
            msg = st.text_input("Enter the text that you received/ the text that you want to check")
            if st.button("Process"):
                data = [msg]
                vec = cv.transform(data).toarray()
                result = model.predict(vec)
                if result[0] == 0:
                    st.success("This is Not A Spam Email")
                    speak("This is Not A Spam Email")
                else:
                    st.error("This is A Spam Email, be aware of viewing this messages")
                    speak("This is A Spam Email, be aware of viewing this messages")
        
            # Add some spacing for better separation
            st.write(" ")

            # Styled text link for feedback with pointing finger emoji
            st.markdown(
                """
                <div style="text-align: center; margin-top: 40px; font-size: 18px;">
                    👉 Click <a href="https://forms.gle/BEDVaaXSxzHgDz3t6" style="color: #3366ff; text-decoration: underline;">here</a> to give feedback
                </div>
                """, unsafe_allow_html=True
            )

        elif about_clicked:
            about_section()

        # Footer
        footer()

    finally:
        # Ensure to uninitialize the COM library even if an exception occurs
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    # Load your machine learning model and vectorizer
    model = pickle.load(open('spam.pkl', 'rb'))
    cv = pickle.load(open('vectorizer.pkl', 'rb'))

    # Run the main function
    main()
