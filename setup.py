from pathlib import Path
from setuptools import setup, find_packages

version = '1.4.0'

setup(name='tnefparse',
      version=version,
      description="a TNEF decoding library written in Python, without external dependencies",
      long_description=Path('README.rst').read_text() + '\n\n' + Path('HISTORY.rst').read_text(),
      classifiers=[
          'Topic :: Communications :: Email',
          'License :: OSI Approved :: GNU Lesser General Public License v3 (LGPLv3)',
          'Intended Audience :: Developers',
          'Intended Audience :: End Users/Desktop',
          'Programming Language :: Python',
          'Programming Language :: Python :: 3',
          'Programming Language :: Python :: 3.6',
          'Programming Language :: Python :: 3.7',
          'Programming Language :: Python :: 3.8',
          'Programming Language :: Python :: 3.9',
          'Programming Language :: Python :: Implementation :: CPython',
          'Programming Language :: Python :: Implementation :: PyPy',
          'Development Status :: 5 - Production/Stable',
          'Environment :: Console',
      ],
      keywords='TNEF MAPI decoding mail email microsoft',
      author='Petri Savolainen',
      author_email='petri.savolainen@koodaamo.fi',
      url='https://github.com/koodaamo/tnefparse',
      license='LGPL',
      packages=find_packages(exclude=['ez_setup', 'examples', 'tests']),
      extras_require={'optional': ['compressed_rtf', ], },
      include_package_data=True,
      zip_safe=True,
      entry_points={
          'console_scripts': ['tnefparse = tnefparse.cmdline:tnefparse']
      })
