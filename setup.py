from setuptools import setup, find_packages

setup(
    name="doc-restructure",
    version="1.0.0",
    description="Word document structure reorganization tool",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    author="Your Name",
    author_email="your@email.com",
    url="https://github.com/congxxx/doc-restructure",
    packages=find_packages(),
    python_requires=">=3.7",
    install_requires=[
        "python-docx>=0.8.11",
    ],
    entry_points={
        "console_scripts": [
            "doc-restructure=doc_restructure.__main__:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Text Processing :: Markup",
    ],
)
