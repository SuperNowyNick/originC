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
#include <GetNBox.h>
////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////////
// Include your own header files here.


////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

// For each column selected finds maximum value and finds point with half of the
// value right to it. Then calculates the difference
int CalculatePlasmonArmLenght()
{
	
	GETN_BOX(trRoot);
	
	GETN_INTERACTIVE(XScale, "X", "");
	GETN_INTERACTIVE(Input, "Input",  "");
	GETN_INTERACTIVE(Output, "Output", "");
	
	if(!GetNBox(trRoot))
		return 1;
	
	DataRange input;
	DataRange output;
	DataRange xscale;
	
   	input.Create(trRoot.Input.strVal, XVT_DATARANGE, false);
   	output.Create(trRoot.Output.strVal, XVT_DATARANGE, false);
   	xscale.Create(trRoot.XScale.strVal, XVT_DATARANGE, false);
   	
   	//get dataranges lengths and so on
   	int ir1, ic1, ir2, ic2, or1, oc1, or2, oc2, xr1, xc1, xr2, xc2;
   	Worksheet ids, ods, xds;
   	string iname, oname, xname;
    input.GetRange(0, ir1, ic1, ir2, ic2, ids, iname);
    output.GetRange(0, or1, oc1, or2, oc2, ods, oname);
    xscale.GetRange(0, xr1, xc1, xr2, xc2, xds, xname);
   	//printf("\n %d, %d, %d, %d, %s, %s \n", xr1, xc1, xr2, xc2, xds.GetName(), xname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", ir1, ic1, ir2, ic2, ids.GetName(), iname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", or1, oc1, or2, oc2, ods.GetName(), oname);
   
    if(oc2-oc1 < 5)
    {
    	MessageBox(GetWindow(), "Less than 6 columns in output. Select at least 6.", "Error!", MB_ICONEXCLAMATION | MB_OK);
    	return 0;
    }
    
    
    // Check if all rows should be processed
    if(ir2==-1){
    	ir2=ids.GetNumRows();
    }
    if(or2==-1){
    	or2=ods.GetNumRows();
    }
    
    // Process every column
    for(int c=0; c<=ic2-ic1; c++){
    	// Get column name
    	string name = ids.Columns(c+ic1).GetComments();
    	// Find max value
    	float val = ids.Cell(ir1,c+ic1);
    	float xval = xds.Cell(xr1, xc1);
    	//printf("\n %f, %f", val, xval);
    	for(int r=0; r<ir2-ir1; r++){
    		if(val < ids.Cell(r+ir1, c+ic1))
    		{
    			val = ids.Cell(r+ir1, c+ic1);
    			xval = xds.Cell(r+xr1, xc1);
    		}
    	}
    	// Maximum found
    	// Find half value
    	float halfval = val/2;
    	float delta = fabs(ids.Cell(ir1,c+ic1)-halfval);
    	//printf("\n%f", delta);
    	float xhalfval = xds.Cell(xr1, xc1);
    	for(int r1=0; r1<ir2-ir1; r1++){
    		if(delta > fabs(ids.Cell(r1+ir1, c+ic1)-halfval) && ids.Cell(r1+ir1, c+ic1)>0)
    		{
    			delta = fabs(ids.Cell(r1+ir1, c+ic1)-halfval);
    			xhalfval = xds.Cell(r1+xr1, xc1);
    			//printf("\n%f", xhalfval);
    		}
    	}
    	//printf("\n Column %i maxval: %f, %f, halfval: %f, %f, diff: %f", c, val, xval, halfval, xhalfval, abs(xval-xhalfval));
    	ods.SetCell(c+or1, oc1, name); // name
    	ods.SetCell(c+or1, oc1+1, xval);
    	ods.SetCell(c+or1, oc1+2, val);
    	ods.SetCell(c+or1, oc1+3, xhalfval);
    	ods.SetCell(c+or1, oc1+4, halfval);
    	ods.SetCell(c+or1, oc1+5, abs(xval-xhalfval));
    }
    ods.Columns(oc1).SetName("Sample");
    ods.Columns(oc1+1).SetName("Maxium x");
    ods.Columns(oc1+2).SetName("Maxium value");
    ods.Columns(oc1+3).SetName("Half of maximum x");
    ods.Columns(oc1+4).SetName("Half of maxium value");
    ods.Columns(oc1+5).SetName("Difference in x");
	return 0;
}