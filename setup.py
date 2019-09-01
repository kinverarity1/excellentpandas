from setuptools import setup

setup(
    name="excellentpandas",
    py_modules=["excellentpandas"],
    version="0.1.1",
    description="Load pandas DataFrames in Excel",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/kinverarity1/excellentpandas",
    install_requires=("xlwings", "pandas"),
)
