from setuptools import setup, find_packages, Command
import sys, os

version = '1.3.0'

with open('README.rst') as readme_file:
    readme = readme_file.read()

with open('HISTORY.rst') as history_file:
    history = history_file.read()

setup(name='tnefparse',
      version=version,
      description="a TNEF decoding library written in python, without external dependencies",
      long_description=readme + '\n\n' + history,
      classifiers=[
       'Topic :: Communications :: Email',
       'License :: OSI Approved :: GNU Lesser General Public License v3 (LGPLv3)',
       'Intended Audience :: Developers',
       'Intended Audience :: End Users/Desktop',
       'Programming Language :: Python',
       'Programming Language :: Python :: 2',
       'Programming Language :: Python :: 3',
       'Programming Language :: Python :: 2.7',
       'Programming Language :: Python :: 3.5',
       'Programming Language :: Python :: 3.6',
       'Programming Language :: Python :: 3.7',
       'Development Status :: 4 - Beta',
       'Environment :: Console',
      ], 
      keywords='TNEF MAPI decoding mail email microsoft',
      author='Petri Savolainen',
      author_email='petri.savolainen@koodaamo.fi',
      url='https://github.com/koodaamo/tnefparse',
      license='LGPL',
      packages=find_packages(exclude=['ez_setup', 'examples', 'tests']),
      setup_requires=["pytest-runner", "pytest-console-scripts"],
      tests_require = ['pytest', 'coverage', 'compressed_rtf'],
      include_package_data=True,
      zip_safe=True,
      entry_points = {
         'console_scripts': ['tnefparse = tnefparse.cmdline:tnefparse']
      }
)
