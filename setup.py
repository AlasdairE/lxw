from Cython.Build import cythonize
from distutils.core import Extension
from distutils.core import setup

extensions = Extension("*", ["src/*.pyx"])

setup(name='lxw',
      ext_modules=cythonize(extensions))
