from setuptools import setup
import os

# Read the requirements from requirements.txt
def read_requirements():
    with open('requirements.txt', 'r', encoding='utf-8') as f:
        return [line.strip() for line in f.readlines() if line.strip() and not line.startswith('#')]

setup(
    name='PDFConvertor',
    version='1.0.0',
    description='A PDF to Word converter application',
    author='Your Name',
    author_email='your.email@example.com',
    packages=['.'],
    install_requires=read_requirements(),
    entry_points={
        'console_scripts': [
            'pdfconvertor=main:main',
        ],
    },
    python_requires='>=3.6',
)