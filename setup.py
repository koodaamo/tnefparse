from setuptools import setup, find_packages, Command
import sys, os

version = '1.1'

class PyTest(Command):
    user_options = []
    def initialize_options(self):
        pass
    def finalize_options(self):
        pass
    def run(self):
        import sys,subprocess
        errno = subprocess.call([sys.executable, 'runtests.py'])
        raise SystemExit(errno)

setup(name='tnefparse',
      version=version,
      description="a TNEF decoding library written in python, without external dependencies",
      long_description="""\
""",
      classifiers=[], # Get strings from http://pypi.python.org/pypi?%3Aaction=list_classifiers
      keywords='TNEF MAPI decoding mail email microsoft',
      author='Petri Savolainen',
      author_email='petri.savolainen@koodaamo.fi',
      url='https://github.com/koodaamo/tnefparse',
      license='LGPL',
      packages=find_packages(exclude=['ez_setup', 'examples', 'tests']),
      include_package_data=True,
      zip_safe=True,
      install_requires=[
          # -*- Extra requirements: -*-
      ],
      entry_points = {
         'console_scripts': ['tnefparse = tnef.cmdline:tnefparse']
      },
      cmdclass = {'test': PyTest},
      )
