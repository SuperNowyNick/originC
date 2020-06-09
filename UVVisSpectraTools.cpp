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
#include <GetNbox.h>

////////////////////////////////////////////////////////////////////////////////////
// Start your functions here.

///////////////////////////////////////////////////////////////////////////////////
//Adjust Spetra
//
// Takes spectrum data that is multiplied accordingly to make even 
// absorbance for a given wavelength
//
// Input: Columns of absorbances
// Reference: Cell with desired absorbance and wavelength
// Output: Columns of absorbances after adjustment
//
///////////////////////////////////////////////////////////////////////////////////

int AdjustSpectra()
{
	
    double mult = 1.00;
    
    
	GETN_BOX(trRoot);
	
	//data range
	GETN_INTERACTIVE(Input, "Input Columns", "[Book1]Sheet1!A");
	
	//data range
	GETN_INTERACTIVE(Reference, "Reference cell", "[Book1]Sheet1!A");
	
    // string edit box 
    GETN_NUM(Constant, "Multiplier", mult)
	
    // choose a column for output data
    GETN_INTERACTIVE(Output, "Output columns", "[Book1]Sheet1!C")  
	
	//show nbox
    if(!GetNBox(trRoot))
    	return 1;
    
    /*
    printf(trRoot.Constant.strVal);
    printf("\n");
    printf(trRoot.Input.strVal);
    printf("\n");
    printf(trRoot.Reference.strVal);
    printf("\n");
    printf(trRoot.Output.strVal);
    printf("\n");
    */
    
    
    DataRange input;
    DataRange output;
    DataRange reference;
    //input.GetRange(trRoot.Input.strVal, ir1, ic1, ir2, ic2, ids);
    //output.GetRange(trRoot.Output.strVal, or1, oc1, or2, oc2, ods);
    //reference.GetRange(trRoot.Reference.strVal, rr1, rc1, rr2, rc2, rds)!=-1)
   	input.Create(trRoot.Input.strVal, XVT_DATARANGE, false);
   	output.Create(trRoot.Output.strVal, XVT_DATARANGE, false);
   	reference.Create(trRoot.Reference.strVal, XVT_DATARANGE, false);
    
   	//get dataranges lengths and so on
   	int ir1, ic1, ir2, ic2, or1, oc1, or2, oc2, rr1, rc1, rr2, rc2;
   	Worksheet ids, ods, rds;
   	string iname, oname, rname;
    reference.GetRange(0, rr1, rc1, rr2, rc2, rds, rname);
    input.GetRange(0, ir1, ic1, ir2, ic2, ids, iname);
    output.GetRange(0, or1, oc1, or2, oc2, ods, oname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", rr1, rc1, rr2, rc2, rds.GetName(), rname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", ir1, ic1, ir2, ic2, ids.GetName(), iname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", or1, oc1, or2, oc2, ods.GetName(), oname);
    

    int refRow = rr1;
    double refVal = rds.Cell(refRow, rc1);
    //printf("RefValue: %f\n", refVal);
    //printf("Refrence row: %d\n", refRow);
    
    
    if(ir2==-1){
    	ir2=ids.GetNumRows();
    }
    if(or2==-1){
    	or2=ods.GetNumRows();
    }
    //printf("Input %d:\n", ic2-ic1);
    //printf("Ouptut %d:\n", oc2-oc1);
    if(ic2-ic1!=oc2-oc1)
    	return 1;
    //printf("Jazda!");
    for(int c=0; c<=ic2-ic1; c++){
    	//printf("OK!");
    	double orgVal = ids.Cell(refRow, c+ic1);
    	//printf("Column number: %d\n", c+ic1);
    	//printf("Column scale factor: %f\n", refVal/orgVal);
    	for(int r=0; r<ir2-ir1; r++){
    		ods.SetCell(r+or1, c+oc1, ids.Cell(r+ir1, c+ic1)*refVal/orgVal*mult);
    	}
    }
    
	return 0;
}

///////////////////////////////////////////////////////////////////////////////////
// AddSubtractSpectra
//
// Takes spectrum data that is multiplied accordingly to make even 
// absorbance for a given wavelength
//
// Input: Columns of absorbances
// Reference: Cell with desired absorbance and wavelength
// Output: Columns of absorbances after adjustment
//
///////////////////////////////////////////////////////////////////////////////////

int SubtractSpectra()
{    
	GETN_BOX(trRoot);
	
	//data range
	GETN_INTERACTIVE(Input, "Input Columns", "[Book1]Sheet1!A");
	
	//data range
	GETN_INTERACTIVE(Reference, "Subtrahend", "[Book1]Sheet1!A");
	
    // choose a column for output data
    GETN_INTERACTIVE(Output, "Output columns", "[Book1]Sheet1!C")  
	
	//show nbox
    if(!GetNBox(trRoot))
    	return 1;
    
    /*
    printf(trRoot.Input.strVal);
    printf("\n");
    printf(trRoot.Reference.strVal);
    printf("\n");
    printf(trRoot.Output.strVal);
    printf("\n");
    */
    
    
    DataRange input;
    DataRange output;
    DataRange reference;
    //input.GetRange(trRoot.Input.strVal, ir1, ic1, ir2, ic2, ids);
    //output.GetRange(trRoot.Output.strVal, or1, oc1, or2, oc2, ods);
    //reference.GetRange(trRoot.Reference.strVal, rr1, rc1, rr2, rc2, rds)!=-1)
   	input.Create(trRoot.Input.strVal, XVT_DATARANGE, false);
   	output.Create(trRoot.Output.strVal, XVT_DATARANGE, false);
   	reference.Create(trRoot.Reference.strVal, XVT_DATARANGE, false);
    
   	//get dataranges lengths and so on
   	int ir1, ic1, ir2, ic2, or1, oc1, or2, oc2, rr1, rc1, rr2, rc2;
   	Worksheet ids, ods, rds;
   	string iname, oname, rname;
    reference.GetRange(0, rr1, rc1, rr2, rc2, rds, rname);
    input.GetRange(0, ir1, ic1, ir2, ic2, ids, iname);
    output.GetRange(0, or1, oc1, or2, oc2, ods, oname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", rr1, rc1, rr2, rc2, rds.GetName(), rname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", ir1, ic1, ir2, ic2, ids.GetName(), iname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", or1, oc1, or2, oc2, ods.GetName(), oname);
   
    if(ir2==-1){
    	ir2=ids.GetNumRows();
    }
    if(or2==-1){
    	or2=ods.GetNumRows();
    }
    //printf("Input %d:\n", ic2-ic1);
   	//printf("Ouptut %d:\n", oc2-oc1);
    if(ic2-ic1!=oc2-oc1)
    	return 1;
    for(int c=0; c<=ic2-ic1; c++){
    	//printf("Column number: %d\n", c+ic1);
    	ods.Columns(c+oc1).SetComments(ids.Columns(c+ic1).GetComments()+"-"+rds.Columns(rc1).GetComments());
    	for(int r=0; r<ir2-ir1; r++){
    		ods.SetCell(r+or1, c+oc1, ids.Cell(r+ir1, c+ic1)-rds.Cell(r+rr1,rc1));
    	}
    }
    
	return 0;
}

int AdjustBaseline()
{
	GETN_BOX(trRoot);
	
	//data range
	GETN_INTERACTIVE(Input, "Input Columns", "[Book1]Sheet1!A");
	
	//data range
	GETN_INTERACTIVE(Reference, "Reference value", "[Book1]Sheet1!A");
	
    // choose a column for output data
    GETN_INTERACTIVE(Output, "Output columns", "[Book1]Sheet1!C")  
	
	//show nbox
    if(!GetNBox(trRoot))
    	return 1;
    
    /*
    printf(trRoot.Input.strVal);
    printf("\n");
    printf(trRoot.Reference.strVal);
    printf("\n");
    printf(trRoot.Output.strVal);
    printf("\n");
    */
    
    DataRange input;
    DataRange output;
    DataRange reference;
    //input.GetRange(trRoot.Input.strVal, ir1, ic1, ir2, ic2, ids);
    //output.GetRange(trRoot.Output.strVal, or1, oc1, or2, oc2, ods);
    //reference.GetRange(trRoot.Reference.strVal, rr1, rc1, rr2, rc2, rds)!=-1)
   	input.Create(trRoot.Input.strVal, XVT_DATARANGE, false);
   	output.Create(trRoot.Output.strVal, XVT_DATARANGE, false);
   	reference.Create(trRoot.Reference.strVal, XVT_DATARANGE, false);
    
   	//get dataranges lengths and so on
   	int ir1, ic1, ir2, ic2, or1, oc1, or2, oc2, rr1, rc1, rr2, rc2;
   	Worksheet ids, ods, rds;
   	string iname, oname, rname;
    reference.GetRange(0, rr1, rc1, rr2, rc2, rds, rname);
    input.GetRange(0, ir1, ic1, ir2, ic2, ids, iname);
    output.GetRange(0, or1, oc1, or2, oc2, ods, oname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", rr1, rc1, rr2, rc2, rds.GetName(), rname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", ir1, ic1, ir2, ic2, ids.GetName(), iname);
    //printf("\n %d, %d, %d, %d, %s, %s \n", or1, oc1, or2, oc2, ods.GetName(), oname);
   
    if(ir2==-1){
    	ir2=ids.GetNumRows();
    }
    if(or2==-1){
    	or2=ods.GetNumRows();
    }
    //printf("Input %d:\n", ic2-ic1);
   	//printf("Ouptut %d:\n", oc2-oc1);
   	float ref = rds.Cell(rr1,rc1);
    if(ic2-ic1!=oc2-oc1)
    	return 1;
    for(int c=0; c<=ic2-ic1; c++){
    	//printf("Column number: %d\n", c+ic1);
    	float colref = ids.Cell(rr1,c+ic1);
    	for(int r=0; r<ir2-ir1; r++){
    		ods.SetCell(r+or1, c+oc1, ids.Cell(r+ir1, c+ic1)-colref+ref);
    	}
    }
    
	return 0;
    
}
