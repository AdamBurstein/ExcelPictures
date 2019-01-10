# ExcelPictures
Create Excel versions of your pictures

This project is a NetCore-2.1 console application which parses photo content for its pixel data and generates a
Microsoft Excel workbook (.xlsx) of that same pixel data.  This is based on a stand-up comedy routine by Matt Parker,
viewable at https://youtu.be/UBX2QQHlQ_I.

Special thanks to my friend Chuck who write the initial version in Java, challenging me to take on this project in .NET.

Installation:
No idea.  I've only run this out of Visual Studio so far.

Execution:
dotnet <dll file> in=<source file> out=<destination file> scale=<x>

Scale allows you to take one out of every <x> pixels in the vertical and horizontal.  For instance, scale=5 takes 1/25 of the
pixel data.
