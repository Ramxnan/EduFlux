from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from django.contrib.auth.models import User
from django import forms
from django.forms.widgets import PasswordInput, TextInput
from .models import File

class CreateUserForm(UserCreationForm):
    class Meta:
        model = User
        fields = ['email', 'password1', 'password2']

    def save(self, commit=True):
        user = super(CreateUserForm, self).save(commit=False)
        email = self.cleaned_data["email"]
        username = email.split('@')[0]  # Extract username from email
        user.username = username
        if commit:
            user.save()
        return user


class LoginForm(AuthenticationForm):
    username = forms.CharField(widget=TextInput())
    password = forms.CharField(widget=PasswordInput())

    def clean(self):
        cleaned_data = super(LoginForm, self).clean()
        email = cleaned_data.get('username')
        if email and '@' in email:
            username = email.split('@')[0]  # Extract username from email
            cleaned_data['username'] = username
        return cleaned_data

class FileForm(forms.ModelForm):
    class Meta:
        model = File
        fields = ['name', 'file']