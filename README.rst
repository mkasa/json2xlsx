=========
json2xlsx
=========
json2xlsx is a tool to generate MS-Excel Spreadsheets from JSON files.

Installation
------------
Install as general python modules. Briefly, do as follows::

    $ sudo easy_install json2xlsx

You can also use pip::

    $ sudo pip install json2xlsx

If you want to install the latest (likely development) version, then do as follows::

    $ cd some_temporary_dir
    $ git clone git://github.com/mkasa/json2xlsx.git
    $ cd json2xlsx
    $ python setup.py build
    $ sudo python setup.py install

Note that you may encounter an error while installing pyparsing on which json2xlsx
depends. This is probably because pyparsing 1.x runs only on Python 2.x while
pyparsing 2.x runs only on Python 3.x. Currently, json2xlsx declares in the package
that pyparsing 1.x is required, which means that Python 3.x users must install
json2xlsx from GitHub with manual modificatin to setup.py. I do not use Python 3.x
often, so please let me know a workaround.

Simple Example
--------------
Let's begin with a smallest example::

    $ cat test.json
    {"name": "John", "age": 30, "sex": "male"}
    $ cat test.ts
    table { "name"; "age"; "sex"; }
    $ json2xslx test.ts -j test.json -o output.xlsx

This will create an Excel Spreadsheet 'output.xlsx' that contains
a table like this:

+-----+-----+-----+
|name | age | sex |
+-----+-----+-----+
|John | 30  | male|
+-----+-----+-----+

Isn't it super-easy? Here, `test.ts` is a script that defines the table.
Let's call it *table script*.
`-j` option specifies an input JSON file. You can specify as many `-j`
as you wish. `-o` gives the name of the output file.

ad hoc Query Example
--------------------
When `-` is specified for input table script, the standard input is used.
`--open` is specified, the generated xlsx file is opened immediately.
Those two options are useful when you want to craete a xlsx file with
an ad hoc query like this::

    $ json2xlsx - -j test.json -o output.xlsx --open
    table { "name"; "age"; "sex"; }
    ^D
    (MS Excel pops up immediately)

Renaming Columns
----------------
Keys in a JSON file are often not appropriate for display use.
For example, you may want to use "Full Name (Family, Given)" instead of
a JSON key "name". You can use `as` modifiers to do this::

    table {
        "name" as "Full Name (Family, Given);
    }

Saving a Few Types
------------------
If a string literal does not contain any spaces, symbols or special characters,
the double quotations can be omitted. This table script::

    table { "name"; "age"; "sex"; }

is equivalent to::

    table { name; age; sex; }

Delimiter
---------
You can use `,` instead of `;`::

    table { name; age; sex; }

`,` and `;` are interchangable except for specifying coordinates.

Adding Styles
-------------
You can add styles to columns::

  table "Analysis Summary" border thinbottom {
    "file_caption" as "Sample" width 20 align right;
    "nSeqs" as "# of \nscaffolds" align right halign center number "#,#";
    "Min" color "green" align right;
    "_1st_Qu" as "1st quantile" align right number "#,#";
  }

1. `width` specifies the width of the column. The unit is unknown (I do not know), so please refer to the openpyxl document for details (although even I have not yet found the answer there).
2. `align right`, `align center`, `align left` will justify the column.
3. `halign right`, `halign center`, `halign left` will justify the heading.
4. `color` specifies the color of the cell.
5. `number` gives the number style of the cell. This will be described in details later.
6. `border` adds a border to the cell.

Number Style
------------
The number style is presumably an internal string used in MS Excel.
Here are a couple of examples.

+---------------------+---------+-----------------------------------+
| Number Format Style | Example | Description                       |
+---------------------+---------+-----------------------------------+
| `%`                 |  24%    | Percentage                        |
+---------------------+---------+-----------------------------------+
| `#,##`              | 123,456 | Insert ',' every 3 digits         |
+---------------------+---------+-----------------------------------+
| `0.000`             |  12.345 | Three digits after decimal point  |
+---------------------+---------+-----------------------------------+

Grouping
--------
You can group multiple columns. An example table script is here::

    table {
        "name";
        group "personal info" {
            "age",
            "sex";
        }
    }

The generated table will look like this.

+-----+---------------+
|     | personal info |
|     +-------+-------+
|name | age   | sex   |
+-----+-------+-------+
|John | 30    | male  |
+-----+-------+-------+

Nesting is allowed.

Multiple Tables, Multiple Sheets
--------------------------------
You can create multiple tables in a sheet::

    # You can write comments here.
    namesheet "Employee List";
    table { "name", "age", "sex"; }
    load "employee1.json";
    # vskip adds a space of specified lines.
    vskip 1;
    table { "company", "revenue"; }
    # You can add as many files.
    load "company1.json";
    load "company2.json";
    # Create a new sheet. The first sheet is implicitly created so we did not need it.
    newsheet;
    namesheet "Products";
    table { "product", "code", "price"; }
    load "product1.json";
    load "product2.json";

Miscellanous
------------
You can use non-ASCII characters. UTF-8 is the only supported coding.

License
-------
Modified BSD License.

Author
------
Masahiro Kasahara

