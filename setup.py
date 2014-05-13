#!/usr/bin/env python3

from setuptools import setup, find_packages
from vcf2csv import __version__

setup(
    name='vcf2csv',
    version=__version__,
    author='movEAX',
    url='https://github.com/movEAX/vcf2csv',
    license='cc by',
    description='Migration address book contacts from Android to Windows Phone by using conversion vCard to CSV',
    py_modules=['vcf2csv'],
    entry_points = {
        'console_scripts': [
            'vcf2csv = vcf2csv:main'
        ]
    },
    zip_safe=False,
)

