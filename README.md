Thrift VBA
==========

A great addition to [Webxcel](https://github.com/michaelneu/webxcel).
You can (almost) make microservices with Excel!

This actually implements the minimal feature set client-wise, not sure how that
happened.

Supports the following transport(s):

 * TSocket
 * TBufferedTransport
 * TFramedTransport
 * THttpClient

Supports the following protocol(s):

 * TBinaryProtocol
 * TMultiplexedProtocol

No servers are implemented (this is a good thing, it should stay like that).

Types are mapped as follow:

Thrift | VBA
------ | ---------
bool   | Boolean
byte   | Byte
i16    | Integer
i32    | Long
i64    | _TLongLong (See note below)_
double | Double
binary | Byte()
string | String

VBA has no 64-bit integer (except, maybe, for LongPtr, but that's only when
running the 64-bit version of Office), so `TLongLong` gives a very basic signed
64-bit integer implementation. Outside of `Equals`, no arithmetic operations are
supported on it, but it does expose `AsLong` and `AsDouble` property to convert
the value to VBA native types. However do note that neither of these types
support the full value range and will either overflow with `Long` (check
`IsValidLong` to see if the value fits beforehand) or lose precision for large
values with `Double`.
