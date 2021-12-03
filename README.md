# Ruby XLSX Converter Comparison

We recently had a need to look for a Ruby XLSX writer that could be used from JRuby for producing a large XLSX file (approx 10 columns, 1 ~ 3M records). The gem we intended to use holds everything in memory and ran out of space.

So, we started to look at options and compared 2 gems that both support streaming so that all the data is not held in memory:
* [XSLX](https://github.com/felixbuenemann/xlsxtream) - a Ruby implementation of the XLSX writer
* [Apache POI](https://poi.apache.org/index.html) - the JAVA API for Microsoft Documents that we can use from within JRuby

The code in this repository is simple sample code to do a quick evaluation on the relative memory consumption and performance to guide the selection between these two gems.

Note that this is just a simple run-time and memory usage comparison - it is not a statement about the feature richness or completeness of the API, etc. though objectively speaking, Apache POI is more feature rich.

## Usage

Make sure that you have JRuby installed. For running the code that uses `xlsxtream`, install the gem by doing `jruby -S gem install xlsxtream`.

After that, just run the code that you want to try:

```
$ jruby use_poi_stream.rb.rb
$ jruby use_xlsxtream.rb
```

Note that you need to press [Enter] at the end of the program to exit back to the command line. We did this just so that it gave us time to look at the memory consumption before it exited - just in case we missed it earlier.

## Results

The scripts are written to create an XLSX file with the following characteristics:
* 1 million rows
* 20 cells in each row
* Each cell has a random integer in it (up to 4 digits)

We got the following results in casual testing using `jruby 9.2.18.0` running on Windows.

* _XLSXStream_: Approx 41 seconds, approx 2000MB memory
* _Apache POI_: Approx 45 seconds, approx 750MB memory

For us, Apache POI was the better choice at this stage but the scripts are available here for anyone who wants to take a look at
improving it or adding new comparisons, etc.

## Contributing

Raise a pull request if you would like to contribute. Please remember to update the README also if you do so.
