from setuptools import setup, find_packages, Command
import sys, os

version = '1.2.2'

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
      long_description="""\n""",
      classifiers=[
       'Topic :: Communications :: Email',
       'License :: OSI Approved :: GNU Lesser General Public License v3 (LGPLv3)',
       'Intended Audience :: Developers',
       'Intended Audience :: End Users/Desktop',
       'Programming Language :: Python',
       'Programming Language :: Python :: 3',
       'Programming Language :: Python :: 2.6',
       'Programming Language :: Python :: 2.7',
       'Programming Language :: Python :: 3.2',
       'Programming Language :: Python :: 3.3',
       'Development Status :: 4 - Beta',
       'Environment :: Console',
      ], 
      keywords='TNEF MAPI decoding mail email microsoft',
      author='Petri Savolainen',
      author_email='petri.savolainen@koodaamo.fi',
      url='https://github.com/koodaamo/tnefparse',
      license='LGPL',
      packages=find_packages(exclude=['ez_setup', 'examples', 'tests']),
      include_package_data=True,
      zip_safe=True,
      entry_points = {
         'console_scripts': ['tnefparse = tnefparse.cmdline:tnefparse']
      },
      cmdclass = {'test': PyTest},
      )
