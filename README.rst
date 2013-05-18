=========
json2xlsx
=========
A tool to generate xlsx (Excel Spreadsheet) files from JSON files

Installation
------------
Install as general python modules. Briefly, do as follows::

    $ mkdir some_temporary_dir
    $ cd some_temporary_dir
    $ git clone git://github.com/mkasa/json2xlsx.git
    $ cd json2xlsx
    $ sudo python setup.py install

json2xlsx is not (yet?) registered in the public repository,
but hopefully in the future we will be able to install json2xlsx
by `easy_install json2xlsx` or `pip install json2xlsx`.

Simple Example
--------------
Let's begin with a smallest example::

    $ cat test.json
    {"name": "John", "age": 30, "sex": "male"}
    $ cat test.xs
    "name", "age", "sex"
    $ json2xmlx test.xs -j test.json -o output.xlsx

This will create an Excel Spreadsheet 'output.xlsx' that contains
a table like this:

+-----+-----+-----+
|name | age | sex |
+-----+-----+-----+
|John | 30  | male|
+-----+-----+-----+

Isn't it super-easy?

License
-------
Modified BSD License.

Author
------
Masahiro Kasahara

