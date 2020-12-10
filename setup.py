from cx_Freeze import setup, Executable

setup(
    name = "myProtimeExtraction",
    version = "0.1",
    description = "Extraction of data of myProtime and save it into excel sheet",
    executables = [Executable("myprotimeextraction_2020.py")]
)