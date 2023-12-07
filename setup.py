from setuptools import setup, find_packages

setup(
    name='pydmcom',
    version='0.1',
    author='dongliangHou',
    author_email='darrenhou1993@gmail.com',
    packages=find_packages(),
    install_requires=[
        'pywin32',
    ],
    description='the python wrapper for the dm.dll',
    url='https://github.com/DarrenHou1993/pydmcom'
)
