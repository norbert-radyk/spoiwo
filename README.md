SPOIWO (Scala POI Wrapping Objects)
==============

[![Build Status](https://travis-ci.org/norbert-radyk/spoiwo.svg?branch=master)](https://travis-ci.org/norbert-radyk/spoiwo)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.norbitltd/spoiwo_2.12/badge.svg)](https://search.maven.org/#search%7Cga%7C1%7Cspoiwo)
[![Javadocs](https://www.javadoc.io/badge/com.norbitltd/spoiwo_2.12.svg)](https://www.javadoc.io/doc/com.norbitltd/spoiwo_2.12)

### Overview

Spoiwo is an open-source library for functional-style spreadsheet generation in Scala. It was started as a wrapper over Apache POI and while the XLSX generation is still at its core, the library has been rectified to support number of other spreadsheet representations (including CSV and HTML).

The library addresses the issues Scala developers face when using spreadsheet libraries for Java and which are representing a highly non-functional way in which the spreadsheets need to be generated (mutable state, enforced indexes, execution order dependency). To address this issues SPOIWO introduces its own spreadsheet model with the number of wrapping classes and caches enabling an efficient report generation. 

This documentation is intended for Spoiwo users and developers to give both an overview and in-depth information of the offered functionality and what problems Spoiwo is intended to solve.

### Download

Spoiwo is available in [The Central Repository](https://search.maven.org/#search%7Cga%7C1%7Cspoiwo).

```
libraryDependencies ++= Seq(
  "com.norbitltd" %% "spoiwo" % "1.6.2"
)
```

### Quick links

* **[About Spoiwo](https://github.com/norbert-radyk/spoiwo/wiki/Spoiwo)** - Introduces the reasons behind creating Spoiwo and presents overview of the core functionality
* **[Quick Start Guide](https://github.com/norbert-radyk/spoiwo/wiki/Quick-start-guide)** - A step-by-step guide to setting up Spoiwo and creating first few simple spreadsheets
* **[Technical Documentation](https://github.com/norbert-radyk/spoiwo/wiki/Technical-documentation)** - Detailed technical documentation on Spoiwo intended for advanced users and project contributors
* **[Release Notes](https://github.com/norbert-radyk/spoiwo/wiki/Release-Notes)** - Overview of the past and coming Spoiwo releases

### Questions or need help

Please check out our [get in touch](https://github.com/norbert-radyk/spoiwo/wiki/Get-In-Touch) page for a different ways of contacting spoiwo team.