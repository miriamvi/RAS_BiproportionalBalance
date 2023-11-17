# RAS. Biproportional Balance method
### The RAS method is a well-known method for data reconciliation. It is an iterative process that adjusts rows and columns back and forth, over and over. It aims to achieve consistency between the entries of some nonnegative matrix and pre-specified row and column totals.  
### This VBA macro performs the biproportional fitting method, which was programmed following the methodology described in Blair and Miller (2009).  
*I developed an Excel VBA application, which works for any matrix regardless of the number of sectors, assuming they have an **m** rows x **n** columns matrix*.  
Instructions:  
1. Open the file **ras-margin.xlsm** and verify that macros are enabled.  
2. Paste the base matrix (the one to be updated through the RAS method) without labels in rows or columns into the first cell (A1). Therefore, only a matrix of exactly **mxn** will be pasted.  
3. In column **n+1**, enter the total value of the target intermediate demand for each sector, i.e., a vector of **mx1** *(It is assumed that, once the ras procedure is done, the sum of the rows for each sector will give you that quantity).*
4. In column **n+2**, enter a formula with the sum of the intermediate demand of the base matrix, which will be different from the target demand recorded in **n+1**.  
5. In row **m+1**, enter the total target value of intermediate purchases for each sector, i.e., a 1xn vector (It is assumed that, once the ras procedure is done, the sum of the columns for the sectors will give you that vector).
6. In row **m+2**, enter the formula with the sum of the intermediate purchases from the base year t.  
7. Observe and memorize the dimension of the matrix (how many rows and columns), excluding the borders.  
8. Click the "RAS" button in the quick access toolbar. If you can't find the button, activate the macro directly in the Developer tab, Macros, RAS0.  
9. a dialog box will open where you must enter the matrix's number of rows and columns.  
10. If the matrix has margins (trade or transport margins), click "Margins," which will open a dialog box to select the margin matrix, which must be of the same dimension as the intermediate transactions matrix. This allows for transitioning from buyer prices to basic prices and updating the matrix.  
If there are no margins, DO NOT click on margins.  
12. Click OK and wait for a new dialog box to appear saying "Generated Matrix."  
