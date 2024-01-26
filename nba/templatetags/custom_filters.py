from django import template

register = template.Library()

@register.filter(name='split_folder_name')
def split_folder_name(value):
    parts = value.rsplit('_', 1)  # Split from the right at the first underscore
    return parts[0] if parts else value
