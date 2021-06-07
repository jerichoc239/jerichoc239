/*********************************************************************************************************
 This script is called by NHForms class in the Pharmacy Windows program to print the customized nursing home record
   It exposes the NHF object that has different printing functions to print e.g. PS, PSSS etc.
   
   
   ###########################	WARNING	 ########################################
   #	ANY Changes Made to This File Can Seriously Affect the Performance 		#
   #	of the Pharmacy Program and can Potentially Make the Program end 		#
   #	Abruptiveley. DO NOT PROCEDE UNLESS YOU ARE SURE OF WHAT YOU ARE DOOING	#
   ##############################################################################
 
        File Created to test and improve Medsheet design .. By Kranthi Sagar
*********************************************************************************************************/

// do not remove the next 3 import lines
import System;
import System.Windows.Forms;
import System.Drawing;

class EventHandlers
{

	 
	static function PrintForm()
	{
		var printDrugImage : Boolean;
		var printWarnings : Boolean;
		var i : int;
		var rxno : long;
	
		
		var lnX : long;
		var lnY : long;
		var lnR : int;
		var lnPatno : int;
		
		var sig : String;
		var adminType : Boolean;
		var repeatDays : int;
		var dosesPerDay : int;
		var lcBrName : String;
		var indexofsig : int;
		var indexofsig2 : int;
		var indexofsig3 : int;
		var end_length : int;
		
		
		var lnX_Offset : long;
		var lnY_Offset : long;
		
		lnR = NHF.mRowno;
		lnPatno = NHF.dtNHData.Rows[lnR]["PATIENTNO"];
		
		if (lnPatno != NHF.mLastPatientNo)
                {		
			NHF.mLastPatientNo=lnPatno;
                        NHF.mPatientPageNo=1;
		}
		else
		{
			NHF.mPatientPageNo++;
		}
//NHF.PSSS(NHF.mPatientPageNo,400,200,8,"B");
		
	/*  Drug image and Warning codes printing */	
		
		printDrugImage = false; // Set this true to print Image and False to remove the image
		printWarnings = true;	// Set this to true to print Warning Codes
		
		lnX_Offset = 0;
		lnY_Offset = 25;
		
		lnX=200-30;
        	lnY=30-10;
		
		lnX += lnX_Offset;
		lnY += lnY_Offset;

		//NHF.PSSS(NHF.oFacility.FACILITYNM.Trim(),lnX+200-100,lnY-30,8,"B");
//NHF.PSSS(NHF.mPatientPageNo,lnX+400,lnY,8,"B");
				
	//	NHF.PSSS("PAKTIA HEALTH MART PHARMACY       ",lnX+800,lnY,9,"B");

		var lnRxC : int;
                var lnLastRowUsed : int;
                lnLastRowUsed=lnR;
		var sPatName : String;
          // NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["FACILITYCD"].ToString().Trim(),lnX,lnY-30,8,"B");
                sPatName=NHF.dtNHData.Rows[lnLastRowUsed]["LNAME"].ToString().Trim()+", "+NHF.dtNHData.Rows[lnLastRowUsed]["FNAME"].ToString().Trim()
             //   NHF.PSSS(sPatName,lnX+1600,lnY-20,8,"B");
	
		lnRxC=0;
NHF.PS(NHF.GetSortOption,lnX+320, lnY+55-30);
		while (true)
		{
			lnRxC++;
                        if (lnRxC > 7)
				break;
			
			if (NHF.dtNHData.Rows[lnR]["PATIENTNO"] != lnPatno)
				break;
				if (NHF.ISPageBreak())
				break;
			if(lnRxC==1)
			{ 
				lnX=80;
				lnY=130+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-30,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				var iRxadminLength: int;
				iRxadminLength = NHF.RxAdminArrayLength;
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 80)			
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	
											
          
				
			}
			else if (lnRxC==2)
			{
				lnX=80;
				lnY=380+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-30,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			else if (lnRxC==3)
			{
				lnX=80;
				lnY=630+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-30,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			else if (lnRxC==4)
			{
				lnX=80;
				lnY=880+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-30,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			else if (lnRxC==5)
			{
				lnX=80;
				lnY=1130+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-30,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			else if (lnRxC==6)
			{
				lnX=80;
				lnY=1380+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-30,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			else if (lnRxC==7)
			{
				lnX=80;
				lnY=1630+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-30,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			
			else if (lnRxC==8)
			{
				lnX=80;
				lnY=1880+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				//PrintRxInfo2(lnX+65,lnY,lnR,printDrugImage);
				Info(lnX+65+65+30,lnY);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				/* if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 80)
				{
				lnRxC++;
				} */
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			else if (lnRxC==9)
			{
				lnX=80;
				lnY=2200+30-30;
				
				lnX += lnX_Offset;
				lnY += lnY_Offset;
				
				PrintRxInfo2(lnX+65,lnY-20,lnR,printDrugImage);
				PrintRxAdmin(lnX+680+150-30+300,lnY+40-15,lnR,printWarnings);
				iRxadminLength = NHF.RxAdminArrayLength;
				/* if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 80)
				{
				lnRxC++;
				} */
				//if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
				
				try
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175||NHF.dtNHData.Rows[lnR]["ExtProfileComment"].Length > 175)
					{
						lnRxC++;
					}
				}
				catch(err) 
				{
					if(iRxadminLength>4||NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 175)
					{
						lnRxC++;
					}
				}	

			}
			
			
			
			 lnLastRowUsed=lnR;
                        lnR++;
                        NHF.mRowno++;
			
			if (NHF.mRowno >= NHF.dtNHData.Rows.Count)
				break;		
		
		
		}
		
		lnX=125;
		lnY=2630-90-90-90-30;
		
		lnX += lnX_Offset;
		lnY += lnY_Offset;
		Info(lnX+65+65+30,lnY-400);
		PrintPatInfo(lnX, lnY, lnLastRowUsed)
		
		
		if (NHF.mRowno >= NHF.dtNHData.Rows.Count)
                        NHF.PR.HasMorePages=false;
                else
                        NHF.PR.HasMorePages=true;
	    	
	    return; 
		
	}
	
	static function PrintDrugInfo(lnX : long, lnY : long, lnR : int,printDrugImage : boolean)
    {
		/*var IsNonRx: Boolean;
		IsNonRx = NHF.dtNHData.Rows[lnR]["IsNonRx"];
		if(IsNonRx)
		{
			PrintNonRxInfo(lnX,lnY,lnR,printDrugImage);
		}
		else
		{	*/
			PrintRxInfo(lnX,lnY,lnR,printDrugImage);
		//}
	}
	
	static function PrintRxInfo(lnX : long, lnY : long, lnR : int,printDrugImage : boolean)
    {
		var xoffset : long;
		xoffset = 0;
		//lnY-=30;
/*
		NHF.PSSS("RX#: "+NHF.dtNHData.Rows[lnR]["rxno"],lnX,lnY,7,"R");
	NHF.PSSS(NHF.dtNHData.Rows[lnR]["DRGNAME"].ToString().Trim()+" "+NHF.dtNHData.Rows[lnR]["STRONG"].ToString(),lnX+200,lnY,7,"B");

				lnY+=20;
				 NHF.PSSS("Qty:"+NHF.dtNHData.Rows[lnR]["QUANT"].ToString()+"  "+"Dr."+NHF.dtNHData.Rows[lnR]["PRESFNM"].Substring(0,1).ToString()+". "+NHF.dtNHData.Rows[lnR]["PRESLNM"].ToString().Trim(),lnX,lnY+10,7,"R");

				lnY+=40;
				if(printDrugImage)
				{
					PrintDrugImage(lnX,lnY,lnR);
					xoffset = 110;
				}
                        	NHF.SF(NHF.CFont,7,"R");
                   if(NHF.dtNHData.Rows[lnR]["SIGLINES"].Length > 80)
{
NHF.SF(NHF.CFont,7,"R");
NHF.PrintRect(NHF.dtNHData.Rows[lnR]["SIGLINES"].ToString(),NHF.CFont, NHF.CBrush,lnX+xoffset,lnY,550,100);
}
else
{       
NHF.SF(NHF.CFont,7,"R"); 	
NHF.PrintRect(NHF.dtNHData.Rows[lnR]["SIGLINES"].ToString(),NHF.CFont, NHF.CBrush,lnX+xoffset,lnY,550,100);
}				
		NHF.PSSS("D/W:"+NHF.dtNHData.Rows[lnR]["DATEO"].ToShortDateString(),lnX,lnY+100,6,"B");			
if (NHF.dtNHData.Rows[lnR]["BRAND"].ToString() != "Y" && NHF.dtNHData.Rows[lnR]["DRGBRNAME"].ToString() != NHF.dtNHData.Rows[lnR]["DRGGENNAME"].ToString())
        	    		{	
					 NHF.PSSS("(GENERIC OF: "+NHF.dtNHData.Rows[lnR]["DRGBRNAME"].ToString().Trim()+")",lnX,lnY+100+30,6,"B");
					
				}
	       			 else
				{
    	         			 NHF.PSSS("(BRAND OF: "+NHF.dtNHData.Rows[lnR]["DRGGENNAME"].ToString().Trim()+")",lnX,lnY+100+30,6,"B");
				}
	*/
	lnX+=350;
	NHF.PSSS(NHF.dtNHData.Rows[lnR]["DRGNAME"].ToString().Trim()+" "+NHF.dtNHData.Rows[lnR]["STRONG"].ToString(),lnX+200,lnY,7,"B");
	lnY+=30;
	NHF.PSSS(NHF.dtNHData.Rows[lnR]["DATEO"].ToShortDateString(),lnX-120-15,lnY-35,7,"B");
	NHF.SF(NHF.CFont,7,"R");
NHF.PrintRect(NHF.dtNHData.Rows[lnR]["SIGLINES"].ToString(),NHF.CFont, NHF.CBrush,lnX,lnY,550,100);
lnY+=90;
//NHF.PSSS(NHF.dtNHData.Rows[lnR]["DATEF"].AddDays(365).ToShortDateString(),lnX-120,lnY-35,7,"B");
if (NHF.dtNHData.Rows[lnR]["BRAND"].ToString() != "Y" && NHF.dtNHData.Rows[lnR]["DRGBRNAME"].ToString() != NHF.dtNHData.Rows[lnR]["DRGGENNAME"].ToString())
        	    		{	
					 NHF.PSSS("RX#: "+NHF.dtNHData.Rows[lnR]["rxno"]+"(GENERIC OF: "+NHF.dtNHData.Rows[lnR]["DRGBRNAME"].ToString().Trim()+")",lnX-120,lnY,6,"B");
					
				}
	       			 else
				{
    	         			 NHF.PSSS("RX#: "+NHF.dtNHData.Rows[lnR]["rxno"]+"(BRAND OF: "+NHF.dtNHData.Rows[lnR]["DRGGENNAME"].ToString().Trim()+")",lnX-120,lnY,6,"B");
				}
	}


static function PrintRxInfo2(lnX : long, lnY : long, lnR : int,printDrugImage : boolean)
    {
		
		lnX+=60;
		NHF.SF(NHF.CFont,9,"B");
		NHF.PS(NHF.dtNHData.Rows[lnR]["DRGNAME"].ToString().Trim()+" "+NHF.dtNHData.Rows[lnR]["STRONG"].ToString().Trim()+" "+NHF.dtNHData.Rows[lnR]["FORM"].ToString(),lnX+60,lnY);
		lnY+=30;
		NHF.SF(NHF.CFont,7,"R");
		//NHF.PSSS(NHF.dtNHData.Rows[lnR]["DATEO"].ToShortDateString(),lnX-120,lnY-35+30,6,"B");
		NHF.SF(NHF.CFont,7,"R"); 
		if(NHF.dtNHData.Columns.Contains("ExtProfileComment") && NHF.dtNHData.Columns.Contains("ExterPat")) 
		{ 
			if(NHF.dtNHData.Rows[lnR]["ExterPat"].ToString().Trim() == "1") 
			{ 
				//NHF.PSSS("Externol profile",lnX,lnY,8,"B"); 
		                       // lnY+=30; 
				//NHF.PrintRect(NHF.dtNHData.Rows[lnR]["ExtProfileComment"].ToString().Trim(),lnX+60+15,lnY+15,550+60+60,100+20+30);
				//lnY+=15; 
				NHF.PrintRect(NHF.dtNHData.Rows[lnR]["ExtProfileComment"].ToString(),NHF.CFont, NHF.CBrush,lnX+60,lnY+5+45,550+60+60,100+20+30+150);
		                       // lnY+=30-15; 
			} 
		}	
		NHF.PrintRect(NHF.dtNHData.Rows[lnR]["SIGLINES"].ToString(),NHF.CFont, NHF.CBrush,lnX+60,lnY+5+45,550+60+60,100+20+30+150);
		lnY+=90+45;

		//NHF.PSSS(NHF.dtNHData.Rows[lnR]["DATEO"].AddDays(365).ToShortDateString(),lnX-120,lnY-35,7,"B");
		lnY+=15;
		if (NHF.dtNHData.Rows[lnR]["BRAND"].ToString() != "Y" && NHF.dtNHData.Rows[lnR]["DRGBRNAME"].ToString() != NHF.dtNHData.Rows[lnR]["DRGGENNAME"].ToString())
      	{
			//	"RX#: "+
			NHF.PSSS(NHF.dtNHData.Rows[lnR]["rxno"],lnX-120,lnY+30,7,"B");
			NHF.PSSS("(GENERIC OF: "+NHF.dtNHData.Rows[lnR]["DRGBRNAME"].ToString().Trim()+")"+"      "+"Dr."+NHF.dtNHData.Rows[lnR]["PRESLNM"].ToString().Trim()+". "+NHF.dtNHData.Rows[lnR]["PRESFNM"].ToString(),lnX-120+200-15,lnY-130,6,"B");		
		}
	    else
		{
			NHF.PSSS(NHF.dtNHData.Rows[lnR]["rxno"],lnX-120,lnY+30,7,"B");
    		NHF.PSSS("(BRAND OF: "+NHF.dtNHData.Rows[lnR]["DRGGENNAME"].ToString().Trim()+")"+"      "+"Dr."+NHF.dtNHData.Rows[lnR]["PRESLNM"].ToString().Trim()+". "+NHF.dtNHData.Rows[lnR]["PRESFNM"].ToString(),lnX-120+200-15,lnY-130,6,"B");
		}
	}
	static function Info(lnX : long, lnY : long)
    	{
        NHF.PSSS("HV- Home Visit, R-Refused, H-Hospitalization ",lnX,lnY,7,"R");
		lnY+=30;
		NHF.PSSS("D/C-Discontinued, P-Given at Program",lnX,lnY,7,"R");
		lnY+=30;		
        NHF.PSSS("X-Medication not ordered to take this time",lnX,lnY,7,"R");
		lnY+=30;
		NHF.PSSS("LPU-Left the Program Unannounced",lnX,lnY,7,"R");
		lnY+=30;
		NHF.PSSS("NA- Not available, WH-Withheld",lnX,lnY,7,"R");
		

    	}
	
	
	static function PrintRxAdmin(lnX : long, lnY : long, lnR : int, printWarnings : boolean)
    {
	var rxno : long;
			 var rxAdminInfo : String[,];
		if(printWarnings)
		{
				var warn1: String = NHF.Warning(lnR,1,0);
				var warn2: String = NHF.Warning(lnR,2,0);
				var warn3: String = NHF.Warning(lnR,3,0);
				var warn4: String = NHF.Warning(lnR,4,0);
				NHF.PSSS(warn1.toLowerCase().Trim()+warn2.toLowerCase().Trim()+warn3.toLowerCase().Trim()+warn4.toLowerCase().Trim(),lnX,lnY-60,7,"B");
		}
		//lnY-=40;
		//lnX+=155;
		rxno=NHF.dtNHData.Rows[lnR]["rxno"];
			
			rxAdminInfo=NHF.GetRxAdminData(rxno);
			
			//getting the lenght of the array for the GetRxAdminData function
			var iRxadminLength: int;
			iRxadminLength = NHF.RxAdminArrayLength;
			
			var yCord=-5;
			
				/*for ( var i=0; i < iRxadminLength;i++)
				{
					if(rxAdminInfo[i,4]=="1")
					{
						//NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
if(rxAdminInfo[i,2].Substring(0,1).length>1)
						{
						NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}
						else
						{					
					NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}					
}
					else if(rxAdminInfo[i,4]=="2")
					{
						yCord=45;
if(rxAdminInfo[i,2].Substring(0,1).length>1)
						{
						NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}
						else
						{					
					NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}
						//NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
					}
					else if(rxAdminInfo[i,4]=="3")
					{
						yCord=90;
if(rxAdminInfo[i,2].Substring(0,1).length>1)
						{
						NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}
						else
						{					
					NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}
						//NHF.PSSS(rxAdminInfo[i,2].Substring(0,1)+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
					}
					else if(rxAdminInfo[i,4]=="4")
					{
						yCord=90+45;
if(rxAdminInfo[i,2].Substring(0,1).length>1)
						{
						NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}
						else
						{					
					NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}
						//NHF.PSSS(rxAdminInfo[i,2].Substring(0,1)+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
					}	
					else
					{
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						yCord+=45;
					}
				}*/
var IsPRN: Boolean;
IsPRN = NHF.dtNHData.Rows[lnR]["ISPRN"];

//NHF.PSSS("LENGTH  "+iRxadminLength.ToString(),lnX+600,lnY,7,"B");
	var yCord=-5;
			if(IsPRN)
			{
				NHF.PSSS("P",lnX+60,lnY+yCord+3,7,"B");
				NHF.PSSS("R",lnX+60,lnY+yCord+3+45,7,"B");
				NHF.PSSS("N",lnX+60,lnY+yCord+3+90,7,"B");

			
			}
			else
			{
				for ( var i=0; i < iRxadminLength;i++)
				{
					if(rxAdminInfo[i,4]=="1")
					{
                                               
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,6,"B");
                           else
{
/*if(rxAdminInfo[i,2].Substring(0,1).length>1)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{					
					NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3+5,7,"B");
}
}		

					else if(rxAdminInfo[i,4]=="2")
					{
						yCord=45;
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3+45,6,"B");
				else
                           {
/*if(rxAdminInfo[i,2].Substring(0,1).length>1)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{					
					NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX-5,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3+5,7,"B");
}
}		else if(rxAdminInfo[i,4]=="3")
					{
						yCord=90;
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,6,"B");
                           else
                         {

/*if(rxAdminInfo[i,2].Substring(0,1).length>2)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{
				//str.split()	
				NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3+5,7,"B");
}
}

					else if(rxAdminInfo[i,4]=="4")
					{
						yCord=90+45;
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,6,"B");
                           else
                         {
/*if(rxAdminInfo[i,2].Substring(0,1).length>1)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{					
				NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3+5,7,"B");
}
}
else if(rxAdminInfo[i,4]=="5")
					{
						yCord=90+45+45+45+30;
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,6,"B");
                           else
                         {
/*if(rxAdminInfo[i,2].Substring(0,1).length>1)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{					
				NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
}
}
else if(rxAdminInfo[i,4]=="6")
					{
						yCord=90+45+90+45+30;
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,6,"B");
                           else
                         {
/*if(rxAdminInfo[i,2].Substring(0,1).length>1)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{					
				NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3+5,7,"B");
}
}	
else if(rxAdminInfo[i,4]=="7")
					{
						yCord=90+45+45+45+45+45+30;
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,6,"B");
                           else
                         {
/*if(rxAdminInfo[i,2].Substring(0,1).length>1)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{					
				NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3+5,7,"B");
}
}	
else if(rxAdminInfo[i,4]=="8")
					{
						yCord=90+45+45+45+45+45+45+30;
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						 if(rxAdminInfo[i,5].Trim()!="")
                         NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,6,"B");
                           else
                         {
/*if(rxAdminInfo[i,2].Substring(0,1).length>1)
{
                         NHF.PSSS(rxAdminInfo[i,2].Substring(0,2).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3,7,"B");
			}
						else
						{					
				NHF.PSSS(rxAdminInfo[i,2].Substring(0,1).Trim()+""+rxAdminInfo[i,2].Substring(4,6),lnX,lnY+yCord+3+5,7,"B");
						}*/
NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3+5,7,"B");
}
}	


	
	
					else
					{
						//NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");
						/* if(rxAdminInfo[i,5].Trim()!="")
                        NHF.PSSS(rxAdminInfo[i,5],lnX,lnY+yCord+3,7,"B");
                           else
                         NHF.PSSS(rxAdminInfo[i,2],lnX,lnY+yCord+3,7,"B");*/
						yCord+=45;
					}

				}



}

			
	}
	
	static function PrintPatInfo(lnX : long, lnY : long, lnLastRowUsed : int)
    {
		var sAllergyLine : String;
		var sAdmitDate : String;
		lnY-=70;
		//lnY-=45;
	//	NHF.PSSS(Convert.ToDateTime(NHF.NHRefFromDate).ToShortDateString(),lnX+200,lnY,8,"B");  
		//	NHF.PSSS(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString(),lnX+900+30,lnY,8,"B");             
		
		lnY+=50;
		
		NHF.PSSS(NHF.GetPrimaryPrescName(lnLastRowUsed),lnX-30,lnY-130,8-1,"B");
			NHF.PSSS(NHF.GetPrimaryPrescPhone(lnLastRowUsed),lnX-30+700,lnY-130,8-1,"B");
		
		if(NHF.dtNHData.Rows[lnLastRowUsed]["PHYSICIAN"].ToString().Trim()=='')
		{
			//NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["PRESLNM"].ToString().Trim()+", "+NHF.dtNHData.Rows[lnLastRowUsed]["PRESFNM"].ToString().Trim(),lnX-30,lnY-130,8-1,"B");
			//NHF.PSSS(NHF.StrFormat(NHF.dtNHData.Rows[lnLastRowUsed]["presphone"],"(XXX)XXX-XXXX").ToString().Trim(),lnX-30+700,lnY-130,8-1,"B");
		}
		else
		{
			NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["PHYSICIAN"].ToString().Trim(),lnX+100,lnY,8,"B");
			NHF.PSSS(NHF.StrFormat(NHF.dtNHData.Rows[lnLastRowUsed]["phone"],"(XXX)XXX-XXXX").ToString().Trim(),lnX+1500+50,lnY,8,"B");	
			//NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["ALTPHYSICIAN"].ToString().Trim(),lnX+100,lnY+30+30,8,"B");
		}
		//NHF.PSSS(Convert.ToDateTime(NHF.NHRefFromDate).ToShortDateString(),lnX+250+350,lnY,7,"B");  
		NHF.SF(NHF.CFont,7,"B");
		//lnY+=95;
                        sAllergyLine = NHF.GetPatAllergyNames(NHF.dtNHData.Rows[lnLastRowUsed]["PATIENTNO"]);
                 // NHF.PSSS(sAllergyLine.PadRight(120,' '),lnX+1300+430,lnY,7,"B");
	NHF.PrintRect(sAllergyLine,lnX+1300+430+15,lnY+15,1000-200,90+10);
	NHF.SF(NHF.CFont,6,"B");
	NHF.PrintRect(NHF.dtNHData.Rows[lnLastRowUsed]["DIET"].ToString().Trim(),lnX+2000+15,lnY-100-60+15+15,1000,90);
	
	lnY+=40+15;
	var sPatName : String;
        sPatName=NHF.dtNHData.Rows[lnLastRowUsed]["LNAME"].ToString().Trim()+", "+NHF.dtNHData.Rows[lnLastRowUsed]["FNAME"].ToString().Trim()
	NHF.PSSS(sPatName.Trim(),lnX+5,lnY,8,"B");
	NHF.PSSS(Convert.ToDateTime(NHF.NHRefFromDate).ToShortDateString(),lnX+2050+500,lnY-8,7+2,"B"); 
        NHF.PSSS("- "+Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString(),lnX+2050+700,lnY-8,8+1,"B"); 
NHF.PSSS(NHF.mPatientPageNo,lnX+3000+50,lnY-8,8+1,"B");
	NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["ROOM"].ToString().Trim(),lnX+1420-820,lnY-8,9,"B");
	NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["BED"].ToString().Trim(),lnX+1245-500,lnY-8,9,"B");
	NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["FACPATNO"].ToString().Trim(),lnX+900,lnY-10,9,"B"); 
if(NHF.dtNHData.Rows[lnLastRowUsed]["ADMITDATE"].ToString().Trim() !="")
			NHF.PSSS(Convert.ToDateTime(NHF.dtNHData.Rows[lnLastRowUsed]["ADMITDATE"]).ToShortDateString(),lnX+2060-840,lnY-8,9,"B");
 
	NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["SEX"].ToString().Trim(),lnX+1450,lnY,8,"B");
	NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["DOB"].ToShortDateString().Trim(),lnX+1520,lnY,8,"B");
	NHF.PSSS("[ "+NHF.oFacility.FACILITYCD.Trim()+" ]",lnX+50,lnY-2400+40,8,"B");
	NHF.PSSS(NHF.oFacility.FACILITYNM.Trim(),lnX+200,lnY-2400+40,8,"B");		
	NHF.SF(NHF.CFont,6,"B");
	NHF.PrintRect(NHF.dtNHData.Rows[lnLastRowUsed]["DIAGNOSIS"].ToString().Trim(),lnX+1000-80+15,lnY-160-40+15,1000,90);
				  lnY+=45;

	/*lnY+=100;
		NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["DIAGNOSIS"].ToString(),lnX+80,lnY+15,8,"B");

		lnY+=90;
		NHF.PSSS(NHF.GetPatInsID(NHF.dtNHData.Rows[lnLastRowUsed]["PATIENTNO"],"MCD").ToString().Trim(),lnX+20,lnY+20,9,"B");
		NHF.PSSS(NHF.GetPatInsID(NHF.dtNHData.Rows[lnLastRowUsed]["PATIENTNO"],"MCR").ToString().Trim(),lnX+600,lnY+20,9,"B");
		lnY+=120+10+15;
		var sPatName : String;
                sPatName=NHF.dtNHData.Rows[lnLastRowUsed]["LNAME"].ToString().Trim()+", "+NHF.dtNHData.Rows[lnLastRowUsed]["FNAME"].ToString().Trim()
                NHF.PSSS(sPatName,lnX+75,lnY,8,"B");
		
		NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["DOB"].ToShortDateString().Trim(),lnX+950,lnY-5,9,"B");
                NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["SEX"].ToString().Trim(),lnX+1245,lnY-8,9,"B");
	        
		NHF.PSSS(NHF.dtNHData.Rows[lnLastRowUsed]["FACPATNO"].ToString().Trim(),lnX+1780,lnY-10,9,"B"); 
		if(NHF.dtNHData.Rows[lnLastRowUsed]["ADMITDATE"].ToString().Trim() !="")
			NHF.PSSS(Convert.ToDateTime(NHF.dtNHData.Rows[lnLastRowUsed]["ADMITDATE"]).ToShortDateString(),lnX+2060-100,lnY-8,9,"B");
	*/
	}
	
	static function FromAndToDate2(lnX : long,lnY : long, size : int) 
	{
		if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==1 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)==2)
				NHF.PSSS("December",lnX,lnY,size,"B");  
			else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==1 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)==1)
				NHF.PSSS("November",lnX,lnY,size,"B");  
			else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==1 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)==0)
				NHF.PSSS("October",lnX,lnY,size,"B"); 
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==9 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("September",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==8 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("August",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==7 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("July",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==6 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("June",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==5 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("May",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==4 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("April",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==3 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("March",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==2 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("February",lnX,lnY,size,"B");  
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(0,1)==1 && Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(1,2)=='/')
				NHF.PSSS("January",lnX,lnY,size,"B");  
				lnX=lnX+250;
				if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().length==10)
					NHF.PSSS(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(6,10),lnX,lnY,size,"B"); 
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().length==9)
					NHF.PSSS(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(5,9),lnX,lnY,size,"B"); 
				else if(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().length==8)
					NHF.PSSS(Convert.ToDateTime(NHF.NHRefToDate).ToShortDateString().substring(4,8),lnX,lnY,size,"B"); 
	}
	
	static function PrintDrugImage(lnX : long, lnY : long, lnR : int)
    {
		var desc: String[];
				desc = NHF.GetDrugDesc(NHF.dtNHData.Rows[lnR]["NDC"]);
				var img: String;
				img=NHF.PrintDrugImage(NHF.dtNHData.Rows[lnR]["NDC"],lnX,lnY,100,70);
				if(img=='true')
				{
					NHF.PrintDrugImage(NHF.dtNHData.Rows[lnR]["NDC"],lnX,lnY,100,70);
				}
				else if(img=='false')
				{
					lnY-=20;
					if(desc[0].length > 1)
					{
						NHF.PSSS("Shp: "+desc[0],lnX,lnY,4,"B");
					}
					if(desc[1].length > 1)
					{
						NHF.PSSS("CF: "+desc[1],lnX,lnY+20,4,"B");
					}
					if(desc[2].length > 1)
					{
						NHF.PSSS("CB: "+desc[2],lnX,lnY+40,4,"B");
					}
					if(desc[3].length > 1)
					{
						NHF.PSSS("ID_F: "+desc[3],lnX,lnY+60,4,"B");
					}
					if(desc[4].length > 1)
					{
						NHF.PSSS("ID_B: "+desc[4],lnX,lnY+80,4,"B");
					}
					if(desc[5].length > 1)
					{
					//	NHF.PSSS("Desc: "+desc[5],lnX,lnY+100,4,"B");
					}
				}

    }

	
}
// please do not remove the next line
NHFrm.add_PrintNHForm(EventHandlers.PrintForm);
