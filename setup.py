# coding: utf-8

from cx_Freeze import setup, Executable

executables = [Executable('main.py')]

setup(name='CreateCompanyCBsUsersFromAnketa',
      version='0.1',
      description='App for Company, Collegial Body and User creating',
      executables=executables)