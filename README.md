SPOIWO (Scala POI Wrapping Objects)
==============

Spoiwo library is a wrapper over Apache POI library dedicated for Scala users supporting functional-style genration of the Excel spreadsheets (only XLSX format).

The aim of the library is to address the ugliness and issues coming from highly non-functional way in which the Excel spreadsheets have to be generated using Apache POI.

To address this issue SPOIWO introduces the number of wrapping classes and caches enabling an efficient report generation. 
Moreover it presents a slightly different approach to the report generation (vs POI) by hiding the row/column indexes from the user 
(as in most situations it really nakes sense to specify empty rows/columns explicitly).

The library is still in the development phase, but hopefully at some point in the near future it will become a nice solution to the issues raised here:
http://stackoverflow.com/questions/5032101/is-there-a-scala-wrapper-for-apache-poi
