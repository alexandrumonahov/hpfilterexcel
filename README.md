# The HP Filter Macro Workbook

The HP Filter Macro Workbook is a collection of functions that calculate the one-sided and two-sided Hodrick-Prescott filters.
The functions are availbale as both a VBA .bas file for developers, as well an easy-to-use Excel macro workbook with examples.

The latest version of the Excel workbook can be found in the [Releases](https://github.com/alexandrumonahov/hpfilterexcel/releases/tag/hp) section.

## HPONE - The One-Sided Hodrick-Prescott Filter

**HPONE(data, lambda, direction)**

  data - the input data to be transformed
  
  lambda - the smoothing parameter
  
  direction - specifies whether the data is oriented vertically (from top to bottom) or horizontally (from left to right)

## HPTWO - The Two-Sided Hodrick-Prescott Filter

**HPTWO(data, lambda, direction)**

  data - the input data to be transformed
  
  lambda - the smoothing parameter
  
  direction - specifies whether the data is oriented vertically (from top to bottom) or horizontally (from left to right)

## Examples:

>=HPONE(A2:A10, 400000)
>
>=HPTWO(A2:M2, 1600, "horizontal")

Since this is an array formula, in Excel versions prior to 2019, users should first select an array of the same size as the
data to be transformed, then enter the function and arguments, and finally press Ctrl + Shift + Enter.

## About

This version of the HP Filter macro was written by Alexandru Monahov

It builds upon the original filters and add-on developed by Kurt Annen

This new version has several improvements in functionality:
1) It extends to the one-sided HP filter the ability to process several series at the same time. Previously, this functionality
   was only available in the two-sided HP filter macro implementation.
2) It allows users to process data which is structured both vertically (from top to bottom), as well as horizontally (from
   left to right), by toggling a newly-implemented 'direction' option.
3) The macro workbook can be launched easily in later versions of Office which limit the usage of the original add-on
   to a single session.

