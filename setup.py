from setuptools import setup, find_packages

setup(
    name="pdf2excel",
    version="0.1.0",
    author="Mehwish",
    author_email="39596277+mehwishh247@users.noreply.github.com",
    description="Convert data tables inside any PDF into Excel sheet",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        "numpy"
        "pytesseract"
    ],
    classifiers=[
        "Programming Language :: Python :: 3.6",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
)
