# JumboExcel - Performant and easy to use OpenXml export framework for .NET #

## Core features ##
Uses only managed NET components.

Optimized for heavy-load server environments.

Thread safe, single export operation per thread (progressive version requires external synchronization per document).

Lazy evaluation over an IEnunumberable allows for exporting huge Excel tables (tens of millions of rows) in a server environment without accumulating memory overhead.

Minimalistic API - single line can be used to generate entire Excel table from IEnumerable.

Supports basic formatting and outlining.

## Performance results ##
Computed on a desktop Dell Inspiron Core i7 when exporting a 100k rows and 10 columns table.

~600k cells per second.

~1.67 seconds to export a million of cells.

More importantly, a server, constantly loaded with huge Excel exporting tasks (hundreds of thousands of rows) on 8 cores did not raise memory footprint over few tens of megabytes.

## Motivation ##
This framework appeared while working on a performance problem, where a client was unable to export a file with 15000 of rows in a desktop application. After a few minutes of running the export routine using Excel automation the problem manifested itself as a nasty out of memory exception.

Generating an excel file on server was out of question as nobody ever wanted such a problem to appear in server environment and nobody from the team wanted to deal with possible compatibility problems.

The framework solved the performance problem and allowed to remove the limitations on certain functions, such as exporting more then 100000 rows, or 200000 cells.
Such tasks are andled by the new framework just instantly so nobody eventhinks of such perf limitations again.