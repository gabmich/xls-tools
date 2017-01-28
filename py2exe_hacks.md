For compiling the two apps with py2exe, I had to trick a little bit :

1. Edit a file in your python venv/lib/openpyxl/__init__.py (on windows, venv/Lib/site-packages/openpyxl/__init__.py) - this modification is necessary for compiling with py2app for osx too

```python
# try:
#     here = os.path.abspath(os.path.dirname(__file__))
#     src_file = os.path.join(here, ".constants.json")
#     with open(src_file) as src:
#         constants = json.load(src)
#         __author__ = constants['__author__']
#         __author_email__ = constants["__author_email__"]
#         __license__ = constants["__license__"]
#         __maintainer_email__ = constants["__maintainer_email__"]
#         __url__ = constants["__url__"]
#         __version__ = constants["__version__"]
# except IOError:
#     # packaged
#     pass
```

became :

```python
__author__ = 'See AUTHORS'
__author_email__ = 'eric.gazoni@gmail.com'
__license__ = 'MIT/Expat'
__maintainer_email__ = 'openpyxl-users@googlegroups.com'
__url__ = 'http://openpyxl.readthedocs.org'
__version__ = '2.4.0-a1'
```