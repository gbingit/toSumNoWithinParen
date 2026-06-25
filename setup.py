from setuptools import setup

APP = ["main.py"]
# 这里写上你的所有第三方库
OPTIONS = {
    "argv_emulation": True,
    # 强制打包openpyxl，防止打包后缺失
    "includes": ["openpyxl"],
    # 关闭多余日志，减小包体积
    "optimize": 2,
}
DATA_FILES = []

setup(
    app=APP,
    name="3toSumNo",  # 最终app名称，可自定义
    data_files=DATA_FILES,
    options={"py2app": OPTIONS},
    setup_requires=["py2app"]
)