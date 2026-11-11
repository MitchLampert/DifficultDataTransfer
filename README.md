# Mitchell's Difficult Data Transfer Tool



Code by Mitchell Scott Lampert  Mitch@MitchLampert.net



This is a data transfer tool designed to play a little loose with the rules, to maximize how much data can be pulled from a source that decided not to play by the rules, itself! This program has been tested on database tables with 6 million rows and 80 columns of various types, with all sorts of various data quality issues strewn into them.



Read the main document, **Mitchells Difficult Data Transfer Program Instruction Manual V1.docx**, for details.



NONE of this code was written with any A.I. It was all pieced together based on my own experience working for banks and such.

## 

## The Principles of Difficult Data Transfer

In accordance with Postel's Law ("...be liberal in what you accept."), the following principles are applied to this program:

* The Data Type at the Source Doesn't Matter!
* Every Individual Row in a Failed Batch Deserves a Second Chance!
* Quantity First, Performance Second!



The idea to get in as much data as possible, and document the hoops it had to jump through in order to do so.

## 

## Use Cases

* You create an SSIS package to transfer a lot of data, but there are a lot of skipped rows because of errors. You can run this program to try to get those rows shoved into the destination table instead! At least until the package is altered to handle more errors at the source.
* You do not, yet, possess the knowledge to use other data transfer tools. Configuring this one should be a lot more simple and straightforward than other options.
* There is a LOT of custom code you need to apply to a data transfer task, and you don’t mind the small performance hit that code like this would make.

## 

## The Two Executables

Source code for two different executable files are included:



DifficultDataAutoGenConfig.exe - Builds a configuration file to be run in the main program.



DifficultDataTransfer.exe and - The main program that actually does the data transfer.



## Requirements

All code was written with C#/.NET 10.0. It has been tested to work on Microsoft Windows 10 and 11. This current source uses OLEDB to connect to databases, but this can be changed in code, if you need to.



It does NOT have any third-party dependencies.

