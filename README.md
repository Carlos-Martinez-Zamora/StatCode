# StatCode

1.	Data input
A)	Foldchange calculation
Simply include the Ct data organized by groups by pairs of data: first the amplicon of interest, followed by the housekeeping amplicon. Include the name of the group only in the first of the two columns.  The Ct data from the sample used as reference (usually the calibrator (cal)) must be as the first data replicas in each columna. 

      A          B          C          D
1   Name1                 Name2
2  GOI name    HK name   GOI name    HK name
3    cal        cal        cal        cal
4    cal        cal        cal        cal
5    cal        cal        cal        cal
6   Sample 1a  Sample 1a  Sample 1b  Sample 1b
7   Sample 1a  Sample 1a  Sample 1b  Sample 1b
8   Sample 1a  Sample 1a  Sample 1b  Sample 1b
9   Sample 2a  Sample 2a  Sample 2b  Sample 2b
10  Sample 2a  Sample 2a  Sample 2b  Sample 2b
11  Sample 2a  Sample 2a  Sample 2b  Sample 2b
  	
When writing the group name, it is important to write it only in the first of the two columns of the group. 

B)	Import data
If you do not wish to calculate the foldchange value from raw Ct data, you may still create a graph and/or statistical analysis by including the data in columns in the same way that foldchange values are exported. Simply add the data from each group in different columns, with the group name as the same line 

       A          B          C          
1   Name1       Name2      Name3
2  Value 1a   Value 1a   Value 1a
3  Value 2a   Value 2a   Value 2a
4  Value 3a   Value 3a   Value 3a
5  Value 4a   Value 4a   Value 4a


If there is any information you do not wish to add or write (like group names), you must still maintain this format. Do not write information in a different place from the one indicated in the example. 
