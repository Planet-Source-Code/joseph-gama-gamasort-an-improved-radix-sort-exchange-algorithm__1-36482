<div align="center">

## Gamasort, an improved Radix Sort Exchange Algorithm


</div>

### Description

* This is a Radix Sort Exchange routine which will work only on

* unsigned integers (unlike a general sorting function which

* works based on comparisons). Since its complexity overhead

* is pretty large, it is efficient on big array where the

* maximum number is bounded by a relatively small ceiling.

* E.g. a big array of bytes or shorts.
 
### More Info
 
* The algorithm is pretty simple:

* 1) It will first find the maximum and minimum radices in order to see

*  if it can skip some radices. Then it proceeds to order

*  from the numbers with the most-significant set-bit downwards

* 2) It leaves the unsorted segments of length == CUTOFF unsorted.

* 3) From higher bit to lower it will compare numbers

*  based on the bit being scanned. This process is

*  improved by using 2 pivots that will move 0's to the

*  left and 1's to the right.

* 4) When finding the pivots it looks for the best case

*  of all 0's or 1's to skip more scanning.

* 5) A low overhead Insertion Sort finishes the job.

*

* To make it an O(n*log(n)) algorithm in which the 'n'

* is proportional to the largest number in the sorted array

* requiers commenting      Call InSort(t(), 1, up)

* and also commenting      (up-lo>CUTOFF)AND

* Which is clearly detailed in the code


<span>             |<span>
---                |---
**Submitted On**   |2002-06-30 12:43:48
**By**             |[Joseph Gama](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/joseph-gama.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Gamasort,\_101092722002\.zip](https://github.com/Planet-Source-Code/joseph-gama-gamasort-an-improved-radix-sort-exchange-algorithm__1-36482/archive/master.zip)








