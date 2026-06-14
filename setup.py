#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
底稿整理助手 v1.0 - setup.py
用于打包成可执行文件
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="docx-format-tool",
    version="1.0.0",
    author="ztywudi",
    author_email="ztyAI@hotmail.com",
    description="底稿整理助手 - Word 报告格式调整工具",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ztywudi/docx-format-tool",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: End Users/Desktop",
        "Topic :: Office/Business :: Office Suites",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3 :: Only",
        "Operating System :: Microsoft :: Windows",
        "Operating System :: POSIX :: Linux",
        "Operating System :: MacOS :: MacOS X",
    ],
    python_requires=">=3.8",
    install_requires=[
        "python-docx>=0.8.11",
    ],
    entry_points={
        "console_scripts": [
            "format-tool=format_tool:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": ["*.ftpl", "*.json"],
    },
)
