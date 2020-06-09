/*------------------------------------------------------------------------------*
 * File Name:				 													*
 * Creation: 																	*
 * Purpose: OriginC Source C file												*
 * Copyright (c) ABCD Corp.	2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010		*
 * All Rights Reserved															*
 * 																				*
 * Modification Log:															*
 *------------------------------------------------------------------------------*/
 
////////////////////////////////////////////////////////////////////////////////////
// Including the system header file Origin.h should be sufficient for most Origin
// applications and is recommended. Origin.h includes many of the most common system
// header files and is automatically pre-compiled when Origin runs the first time.
// Programs including Origin.h subsequently compile much more quickly as long as
// the size and number of other included header files is minimized. All NAG header
// files are now included in Origin.h and no longer need be separately included.
//
// Right-click on the line below and select 'Open "Origin.h"' to open the Origin.h
// system header file.
#include <Origin.h>
////////////////////////////////////////////////////////////////////////////////////

//#pragma labtalk(0) // to disable OC functions for LT calling.

////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.


////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.
int XMLImport(Page& pgTarget, TreeNode& trFilter, LPCSTR lpcszFile, int nFile)
{
	if(pgTarget)
	{
		Tree xmlTree;
		xmlTree.Load(lpcszFile, 0);
		
		int colcount = 4;
		int rowcount = xmlTree.Worksheet.Table.GetNodeCount()-4;
		
		vector<string> treeData;
		
		tree_get_values(xmlTree.Worksheet.Table, treeData, false);
		
		Worksheet wks = pgTarget.Layers(0);				// Get the active Worksheet from Layers collection
		if( wks )                                   // If wks IsValid...
		{
			for(int jj=0; jj<colcount-2; jj++)
			{
				Column col = wks.Columns(jj+2*nFile);			// Get first Column in Worksheet
				if( !col )                               // If col IsInvalid...
				{ 
					col = wks.Columns(wks.AddCol());
				}
				if(jj==0)
					col.SetType(OKDATAOBJ_DESIGNATION_X);
				col.SetLongName(treeData[5+1+(jj*2)+1]);
				col.SetComments(treeData[5+9+1+4+((1-jj)*2)+1]);
				Dataset ds(col);					// Get DataObject (Dataset) of column
				if( ds )                            // If ds IsValid...
				{
					ds.SetSize(rowcount-1);                 // Set size of data set
					for(int ii = 1; ii < rowcount; ii++ )
						ds[ii-1]=atof(treeData[5+(ii*9)+1+(jj*2)+1]);             
				}	
			}
		}
	}

	return 0;
}
