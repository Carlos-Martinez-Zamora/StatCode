# StatCode

The StatCode file must be open (R running) to use this app.

## Data input

The file "StatCode_Template" indicates how the data must be arranged in the **.xlsx file** needed as input. If there is any information you do not wish to add or write (like group names), you must still maintain this format. Do not write information in a different place from the one indicated in the example. When analyzing the same file several times, you must re-load the document with every new iteration in order to not lose the previous results

#### Foldchange calculation
Simply include the Ct data organized by groups by pairs of data: first the amplicon of interest, followed by the housekeeping amplicon. Include the name of the group only in the first of the two columns.  The Ct data from the sample used as reference (usually the calibrator (cal)) must be as the first data replicas in each columna. 

When writing the group name, it is important to write it only in the first of the two columns of the group. 

#### Import data
If you do not wish to calculate the foldchange value from raw Ct data, you may still create a graph and/or statistical analysis by including the data in columns in the same way that foldchange values are exported. Simply add the data from each group in different columns, with the group name as the same line 

## Statistical tests
The statistical analysis follows the following flowchart shown in the file "StatCode_Template". In the case of the normality test, if any group follows a non-normal distribution, the data is transformed using logarythmic, exponential and inverse trnasformations, if allowed by the dataset. 






