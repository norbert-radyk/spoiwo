SPOIWO (Scala POI Wrapping Objects)
==============

Spoiwo library is a wrapper over Apache POI library dedicated for Scala users and supporting functional-style generation of the Excel spreadsheets (only XLSX format).

The aim of the library is to address the ugliness and issues coming from highly non-functional way in which the Excel spreadsheets have to be generated using Apache POI. To address this issue SPOIWO introduces the number of wrapping classes and caches enabling an efficient report generation. 

Moreover it presents a slightly different approach to the report generation (vs POI) by hiding the row/column indexes from the user (as in most situations it really nakes sense to specify empty rows/columns explicitly).

The library is still in the development phase, but hopefully at some point in the near future it will become a nice solution to the issues raised here:
http://stackoverflow.com/questions/5032101/is-there-a-scala-wrapper-for-apache-poi

### Questions or need help

Please check out our [get in touch](https://github.com/norbert-radyk/spoiwo/wiki/GetInTouch) page for a different ways of contacting spoiwo team.

### Copyright and license

Spoiwo is copyright 2014 NorbIT Ltd.

Licensed under [MIT License](http://opensource.org/licenses/MIT) you may not use this software except in compliance with the License.

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
