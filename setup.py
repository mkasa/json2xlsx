#!/usr/bin/env python

from setuptools import setup

setup(name='json2xlsx',
      version='1.0',
      description='Tool to generate xlsx (Excel spreadsheet) from JSON',
      author='Masahiro Kasahara',
      author_email='mkasa@cb.k.u-tokyo.ac.jp',
      url='http://github.com/mkasa/json2xlsx',
      license='BSD',
      packages=['json2xlsx', 'json2xlsx.utilities'],
      zip_safe=True,
      classifiers=[
          'Development Status :: 4 - Beta',
          'Environment :: Console',
          'Intended Audience :: Developers',
          'Intended Audience :: End Users/Desktop',
          'Intended Audience :: Science/Research',
          'License :: OSI Approved :: BSD License',
          'Natural Language :: English',
          'Operating System :: OS Independent',
          'Programming Language :: Python',
          'Programming Language :: Python :: 2.7',
          'Topic :: Scientific/Engineering :: Information Analysis',
          'Topic :: Software Development :: Libraries :: Python Modules',
          'Topic :: Utilities'
      ],
      entry_points= {
          'console_scripts': [
              'json2xlsx = json2xlsx.utilities.json2xlsx:main'
          ]
      },
      install_requires = [
          'argparse>=1.2.1',
          'openpyxl>=1.5.7',
          'pyparsing>=1.5.5'
      ],
      )

