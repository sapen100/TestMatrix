TEST MATRIX APPLICATION

This application will allow you to generate the test matrix automatically based on the bulletins that are received from the VSP program.

HOW TO EXECUTE THE APPLICATION

   1. Downlaod the VSP Bulletins (.zip). Extract all bulletins(*.docx) that are in the zip file (You could use the BulletinUncompressor Tool to do it).
   2. Execute TMApplication.exe
   3. Once the application is opened, you have to select the type of Test Matrix you would like to generate. 
   4. Select Bulletins location, which is the uncompressed folder in step 1.
   5. Press "Start Generating TM" button.


TEMPLATES THAT MUST BE IN THE SAME FOLDER AS TMApplication.exe

   OriginalTM.xlsx This template is being used to generate the Test Matrix. 

   Apps.xlsx       This file contains the default platforms of the application.



EXCEPTIONS

1. We extract the content from the tables that are in the Bulletins. The relevant information for us is the content where the KB resides. However, we ran into the following 
   exception. The relevant content is not in the same Cell where the KB is located. In order to fix this exception, you have to have "Microsoft Office SharePoint Server 2007 Service Pack 2 (32-bit editions)" loaded in the Apps.xlsx   

Microsoft Office 									Maximum 		Aggregate 			Bulletins Replaced 
Suite and Other Software				Component			Security Impact	 	Severity Rating			by this Update
-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Microsoft Office SharePoint Server 			Excel Services[2]		Remote Code 		Important			MS10-017
2007 Service Pack 2 (32-bit editions)			(KB2553093)			Execution	



