# LibreOffice4j

## Overview
This library aims to provide a simple and dev-friendly access to the LibreOffice UNO API.

The following issues with the UNO API are:
- Extremely complicated API.
- Very steep learning curve.
- Something that should be done in one line requires many many more in the UNO API.

This API wraps around the UNO API and provide a much more dev-friendly set of interfaces.  
The API can either provide a full featured API or simply wrap around specific classes of the UNO API - there is no hard requirement.

## Status
Due to the high complexity of the LibreOffice UNO API, not all interfaces/methods will be implemented at once.  
Instead, they will be added as users of this library require different bits.

## Getting started
TODO

## Convention
- K.I.S.S.
- Only unchecked exceptions: All LibreOffice checked exceptions will be encapsulated/adapted.
