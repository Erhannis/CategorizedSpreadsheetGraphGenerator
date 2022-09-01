You know when you're generating multiple kinds of data, so you write lines to a CSV where one column is the kind of data, another is the timestamp, and a third is the value?  You know how then you want to open that CSV in Excel, and make a scatter plot of the data, with one series per category?  ...And you know how Excel doesn't know how to do that???  After running into this multiple times, I made this program that takes a csv, splits categories into separate sheets, and makes a chart of all the data, and outputs an xlsx.

Run it like
```
cat FILE.csv | java -jar CSGG.jar
```

It creates an xlsx in the current directory.

Released under the MIT license.