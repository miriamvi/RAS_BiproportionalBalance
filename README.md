# RAS_BiproportionalBalance
The RAS method is a well-known method for data reconciliation. It is an iterative process that adjusts rows and columns back and forth, over and over.Its aim is to achieve consistency between the entries of some nonnegative matrix and pre-specified row and column totals. 
The macro programming that performs the biproportional fitting method was carried out in accordance with the methodology described in Blair and Miller (2009), detailed in the appendix.
I developed an Excel VBA application, which works for any matrix regardless of the number of sectors, assuming they have an m rows x n columns matrix. 
Application:
Open the file ras-margin.xlsm and verify that macros are enabled.\n
Paste the base matrix (the one to be updated) without labels in rows or columns into the first cell (A1). Therefore, only a matrix of exactly mxn will be pasted.
In column n+1, enter the total value of the target intermediate demand for each sector, i.e., a vector of mx1 (It is assumed that, once the ras procedure is done, the sum of the rows for each sector will give you that quantity).
In column n+2, enter a formula with the sum of the intermediate demand of the base matrix, which will be different from the target demand recorded in n+1.
In row m+1, enter the total target value of intermediate purchases for each sector, i.e., a 1xn vector (It is assumed that, once the ras procedure is done, the sum of the columns for the sectors will give you that vector).
In row m+2, enter the formula with the sum of the intermediate purchases from the base year t.
Observe and memorize the dimension of the matrix (how many rows and columns), excluding the borders.
Click on the "RAS" button located in the quick access toolbar. If you cannot find the button, you can activate the macro directly in the Developer tab, Macros, RAS0.
Next, a dialog box will open where you must enter the matrix's number of rows and columns.
If the matrix has margins, click on "Margins," which will open a dialog box to select the margin matrix, which must be of the same dimension as the intermediate transactions matrix. This allows for the transition from buyer prices to basic prices and updating the matrix.
If there are no margins, DO NOT click on margins.
Click OK and wait for a new dialog box to appear saying "Generated Matrix."
