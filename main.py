import os
import pandas as pd
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.popup import Popup
from kivy.core.window import Window
from kivy.metrics import dp
from kivy.uix.progressbar import ProgressBar
from kivy.clock import Clock
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import socket

class MessageApp(App):
    title = "Fusion Messenger"

    def build(self):
        # Set window properties
        Window.size = (900, 700)
        Window.clearcolor = (0.15, 0.15, 0.2, 1)

        self.main_layout = BoxLayout(orientation='vertical', padding=dp(15), spacing=dp(10))

        # Header
        header = Label(text="Fusion Messenger", size_hint=(1, 0.1), font_size=dp(24), bold=True)
        self.main_layout.add_widget(header)

        content_area = BoxLayout(orientation='horizontal', size_hint=(1, 0.9))
        left_panel = BoxLayout(orientation='vertical', size_hint=(0.4, 1), spacing=dp(10))

        # File chooser
        upload_label = Label(text="Upload Contact File", size_hint=(1, 0.2))
        self.file_chooser = FileChooserListView(path=os.path.expanduser("~"), filters=["*.xlsx", "*.csv"], size_hint=(1, 0.6))
        self.upload_button = Button(text="Upload File", size_hint=(1, 0.2))
        self.upload_button.bind(on_press=self.load_file)

        left_panel.add_widget(upload_label)
        left_panel.add_widget(self.file_chooser)
        left_panel.add_widget(self.upload_button)

        right_panel = BoxLayout(orientation='vertical', size_hint=(0.6, 1), spacing=dp(10))

        # Message input
        composer_label = Label(text="Message Composer", size_hint=(1, 0.1))
        self.message_input = TextInput(hint_text="Type your message here...", size_hint=(1, 0.8), multiline=True)
        self.preview_button = Button(text="Preview Message", size_hint=(1, 0.1))

        right_panel.add_widget(composer_label)
        right_panel.add_widget(self.message_input)
        right_panel.add_widget(self.preview_button)

        # Email credentials and send
        email_creds = BoxLayout(orientation='vertical', size_hint=(0.7, 1))
        self.email_input = TextInput(hint_text="Your Email", size_hint=(1, 0.5), multiline=False)
        self.password_input = TextInput(hint_text="Password/App Password", size_hint=(1, 0.5), multiline=False, password=True)
        email_creds.add_widget(self.email_input)
        email_creds.add_widget(self.password_input)

        self.email_button = Button(text="Send Emails", size_hint=(0.3, 1))
        self.email_button.bind(on_press=self.send_emails)

        email_section = BoxLayout(size_hint=(1, 0.3))
        email_section.add_widget(email_creds)
        email_section.add_widget(self.email_button)

        right_panel.add_widget(email_section)

        content_area.add_widget(left_panel)
        content_area.add_widget(right_panel)

        self.main_layout.add_widget(content_area)
        self.contacts_data = None

        return self.main_layout

    def load_file(self, instance):
        path = self.file_chooser.selection
        if path:
            file_path = path[0]
            try:
                if file_path.endswith('.xlsx'):
                    self.contacts_data = pd.read_excel(file_path)
                else:
                    self.contacts_data = pd.read_csv(file_path)
                self.show_message("File loaded successfully.")
            except Exception as e:
                self.show_message(f"Error loading file: {str(e)}")

    def send_emails(self, instance):
        if self.contacts_data is None or 'Email Address' not in self.contacts_data.columns:
            self.show_message("No contacts with email addresses loaded.")
            return
        if not self.message_input.text.strip():
            self.show_message("Please enter a message first.")
            return
        if not self.email_input.text or not self.password_input.text:
            self.show_message("Please enter your email credentials.")
            return

        progress_popup = self.create_progress_popup("Sending Emails")
        progress_popup.open()
        Clock.schedule_once(lambda dt: self._process_emails(progress_popup), 0.1)

    def _process_emails(self, popup):
        try:
            emails = self.contacts_data['Email Address'].dropna().tolist()
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.email_input.text, self.password_input.text)

            success_count = 0
            for i, email in enumerate(emails):
                msg = MIMEMultipart()
                msg['From'] = self.email_input.text
                msg['To'] = email
                msg['Subject'] = "Message from Fusion Messenger"
                msg.attach(MIMEText(self.message_input.text, 'plain'))
                server.sendmail(self.email_input.text, email, msg.as_string())
                success_count += 1

            server.quit()
            popup.dismiss()
            self.show_message(f"Sent {success_count} of {len(emails)} emails successfully.")

        except Exception as e:
            popup.dismiss()
            self.show_message(f"Error: {str(e)}")

    def create_progress_popup(self, title):
        content = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        label = Label(text="Processing...", size_hint=(1, 0.3))
        progress = ProgressBar(max=100, value=0, size_hint=(1, 0.4))
        button = Button(text="Cancel", size_hint=(1, 0.3))
        content.add_widget(label)
        content.add_widget(progress)
        content.add_widget(button)

        popup = Popup(title=title, content=content, size_hint=(0.8, 0.3), auto_dismiss=False)
        button.bind(on_press=popup.dismiss)
        return popup

    def show_message(self, message):
        content = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        label = Label(text=message, size_hint=(1, 0.7))
        button = Button(text="OK", size_hint=(1, 0.3))
        content.add_widget(label)
        content.add_widget(button)
        popup = Popup(title="Message", content=content, size_hint=(0.8, 0.3))
        button.bind(on_press=popup.dismiss)
        popup.open()

if __name__ == '__main__':
    MessageApp().run()
