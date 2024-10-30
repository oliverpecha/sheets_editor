from setuptools import setup, find_packages

setup(
    name="sheets_editor",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        'gspread',
        'google-auth',
        'pyyaml'
    ],
    author="Your Name",
    author_email="your.email@example.com",
    description="A package for editing and managing Google Sheets",
    long_description=open('README.md').read(),
    long_description_content_type="text/markdown",
    url="https://github.com/oliverpecha/sheets_editor",
)
