from django.core.exceptions import ValidationError

class SimplePasswordValidator:
    def validate(self, password, user=None):
        pass  # Do nothing for validation

    def get_help_text(self):
        return "Your password can be any length and contain any characters."
