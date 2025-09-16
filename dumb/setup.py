#!/usr/bin/env python3
"""
Setup configuration for ZERF Automation System
"""

from setuptools import setup, find_packages
from pathlib import Path

# Read the contents of README file
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding='utf-8')

# Read requirements
requirements = []
with open('requirements.txt') as f:
    for line in f:
        line = line.strip()
        if line and not line.startswith('#'):
            requirements.append(line)

setup(
    name="zerf-automation-system",
    version="2.0.0",
    author="Lam Research Development Team",
    author_email="your.email@lamresearch.com",
    description="SAP ZERF Data Extraction and Processing Automation System",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/your-org/zerf-automation-system",
    package_dir={"": "src"},
    packages=find_packages(where="src"),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: System :: Systems Administration :: Automation",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    extras_require={
        "dev": [
            "pytest>=7.4.0",
            "pytest-mock>=3.11.1",
            "pytest-cov>=4.1.0",
            "black>=23.7.0",
            "flake8>=6.0.0",
        ],
        "gui": [
            "tkcalendar>=1.6.1",
        ],
    },
    entry_points={
        "console_scripts": [
            "zerf-automation=main:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": ["*.ini", "*.json", "*.yaml", "*.yml"],
    },
    zip_safe=False,
    keywords="sap automation excel data-processing sharepoint",
    project_urls={
        "Bug Reports": "https://github.com/your-org/zerf-automation-system/issues",
        "Source": "https://github.com/your-org/zerf-automation-system",
        "Documentation": "https://github.com/your-org/zerf-automation-system/docs",
    },
)