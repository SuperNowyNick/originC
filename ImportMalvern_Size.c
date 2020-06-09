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


////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.


////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

int MalvernGetVars(StringArray& saVarNames, StringArray& saVarValues, const StringArray& saHdrLines, const TreeNode &trFilter);
{
	
}

int MalvernImport(Page& pgTarget, TreeNode& trFilter, LPCSTR lpcszFile, int nFile)
{
	if(pgTarget)
	{
		Page pg = Project.Pages();
		string strPageName = pg.GetName();
		int nIndexLayer = pg.Layers().GetIndex();
		string strPath = GetAppPath(TRUE) +
	}
}
