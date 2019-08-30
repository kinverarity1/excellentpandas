# excellentpandas

[![License](http://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/kinverarity1/excellentpandas/blob/master/LICENSE)

Very quickly load pandas DataFrames in Excel

Python is awesome and I love it for doing all sorts of data manipulation. But sometimes Microsoft Excel remains the best place to do quick data exploration and filtering. So thanks to the brilliant [xlwings](xlwings), it's easy to integrate the two. This module has some very simple functions to make this as easy as possible.

Say you have this script:

```python
>>> df = read_data()
>>> result = df.groupby("blah")[vars].agg(something_complex)
```

And you want to explore ``result`` quickly in Excel, you can try one of the following functions to immediately launch the DataFrame in a new Excel Workbook on your desktop via a non-blocking call:

``show_in_excel(df)``

```python
>>> from excellentpandas import show_in_excel
>>> df = read_data()
>>> result = df.groupby("blah")[vars].agg(something_complex)
>>> show_in_excel(result)
>>>
```

Better - use ``.pipe(via_excel)``:

```python
>>> from excellentpandas import via_excel
>>> df = read_data()
>>> result = df.groupby("blah")[vars].agg(something_complex).pipe(via_excel)
>>>
```

Or, I like to also print info about the DataFrame: use ``.pipe(via_info_excel)``

```python
>>> from excellentpandas import via_info_excel
>>> df = read_data()
>>> result = df.groupby("blah")[vars].agg(something_complex).pipe(via_info_excel)
<class 'pandas.core.frame.DataFrame'>
Int64Index: 16 entries, 2004 to 2019
Data columns (total 6 columns):
...
dtypes: float64(6)
memory usage: 896.0 bytes
>>>
```

## Install

```
$ pip install -U excellentpandas
```
