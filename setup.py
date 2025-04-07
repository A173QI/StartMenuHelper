"""
Start Menu Shortcut Creator - Setup Script
This script is for creating a Windows executable (.exe) file
"""
from setuptools import setup, find_packages

setup(
    name="Start Menu Shortcut Creator",
    version="1.0.0",
    description="A Windows desktop application that simplifies creating Start Menu shortcuts",
    author="Replit",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "PyQt5>=5.15.0",
        "pywin32>=228;platform_system=='Windows'",
    ],
    python_requires=">=3.6",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Topic :: Utilities",
    ],
    entry_points={
        "console_scripts": [
            "start-menu-shortcut-creator=main:main",
        ],
    },
    options={
        "build_exe": {
            "packages": ["os", "sys", "ctypes", "win32com", "win32api", "PyQt5"],
            "include_files": ["assets/"],
        },
    },
)