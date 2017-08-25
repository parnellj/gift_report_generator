try:
	from setuptools import setup
except ImportError:
	from distutils.core import setup

config = {
	'name': 'College of Science Gift Report Generator',
	'version': '0.1',
	'url': 'https://github.com/parnellj/gift_report_generator',
	'download_url': 'https://github.com/parnellj/gift_report_generator',
	'author': 'Justin Parnell',
	'author_email': 'parnell.justin@gmail.com',
	'maintainer': 'Justin Parnell',
	'maintainer_email': 'parnell.justin@gmail.com',
	'classifiers': [],
	'license': 'GNU GPL v3.0',
	'description': 'A tool to rerieve and manipulate gift report files from the Advance donor management system using PyAutoGUI.',
	'long_description': 'A tool to rerieve and manipulate gift report files from the Advance donor management system using PyAutoGUI.',
	'keywords': '',
	'install_requires': ['nose'],
	'packages': ['gift_report_generator'],
	'scripts': []
}
	
setup(**config)
