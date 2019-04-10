from setuptools import setup

setup(name='brotherprinter',
      version='1.0',
      description='Python module for printing to Brother Printers via b-Pac SDK',
      url='https://github.com/JonkopingUniversityLibrary/brotherprinter',
      author='Gustav Lindqvist',
      author_email='gustav.lindqvist@ju.se',
      license='Unlicense',
      packages=['brotherprinter'],
      install_requires=[
            'pypiwin32'
      ],
      include_package_data=True,
      zip_safe=False)