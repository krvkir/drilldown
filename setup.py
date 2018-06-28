#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""The setup script."""

from setuptools import setup, find_packages

with open('README.org') as readme_file:
    readme = readme_file.read()

requirements = ['pandas', 'xlsxwriter', ]

setup_requirements = ['pytest-runner', ]

test_requirements = ['pytest', ]

setup(
    author="Kirill Krasnoshchekov",
    author_email='krvkir@gmail.com',
    description="Generates xlsx reports with hierarchically organized sheets.",
    install_requires=requirements,
    license="GNU General Public License v3",
    long_description=readme
    include_package_data=True,
    keywords='drilldown',
    name='drilldown',
    packages=find_packages(include=['drilldown']),
    setup_requires=setup_requirements,
    test_suite='tests',
    tests_require=test_requirements,
    url='https://github.com/krvkir/drilldown',
    version='0.1.0',
    zip_safe=False)
