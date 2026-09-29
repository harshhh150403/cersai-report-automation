**Batch size will be equal to total unique collateral ids**   
**Date format everywhere is yyyy-mm-dd**  
**We need to create an excel output which will incrementally work on each collateral id. Of the batch and create their rows in the below way-\>**  
**We will only be working on these rows(sheets in CERSAI2\_Registartion\_BatchFormatDescription excel)- introduction, is, immovable\_asset, borrower, charge\_holder**

**Row 1 (this row will be written only once in output excel for all batches)**

 **Column 1 \- File header \- value(“FH”)**  
 **Column 2 \- File Type \- value(“SI\_REGN”)**  
 **Column 3 \- Batch size**  
 **column 4 \-Today’s Date**

  **(from here-\> each collateral of batch will get all these rows)**  
 **Row 2**  
**Column 1 \- File header \- value(“SI”)**  
 **Column 2 \- Serial number(unique value given to a particular unique collateral no. of a batch) for eg \- if batch size is N- (serial no. will have value \- 1 to N)**  
 **Column 3 \- total no. of borrowers (get value from counting  unique Neo CIF ID of the collateral id. in collateral excel)**  
 **Column 4 \- in collateral excel \- in the collateral id which we are working in, find all unique “Owner Name” and find how many of them are also listed in Customer Name in any row of that particular collateral id. If their name is present in customer name, don’t increase the count. Write total count of Owners who are not borrower in this column.**  
**Column 5 \- 0 by default**  
 **Column 6 \- SI\_type\_ID  (get “type of mortgage”from req document of this application id. no-\> then in SI\_SUB\_TYPE\_ID sheet in CERSAI2.0MASTERS excel, find the type of mortgage in “SI Type Description” column and get its corresponding “SI Sub Type Id / Transaction Sub Type ID” value and fill this value in this column)**  
 **Column 7 \- by default (blank space)**  
 **Column 8 \- by default financing type is “sole” and its id is 1, so by default \- value(1)**   
**Column 9 \- disbursal date from collateral excel of that application id.**  
 **Column 10 \- Total\_secured\_amount \- get value from \-\> (get sanctioned loan amount from any one of the rows of collateral id which we are working on in (“collateral” excel table))**  
 **Column 11 \-  “Application id” \+ “Collateral id” (concat) from  collateral excel**  
 **Column 12 \- by default (“NA”)**  
**Column 13 \-  by default (“FIRST CHARGE”)**  
 **Column 14 \- by default (0)**  
**Column 15 \-  BATCH\_{current\_date} yyy-mm-dd**

 **Row 3**   
**Column 1 \- File header \- value(“IMM”)**  
 **Column 2 \- by default (1)**  
 **Column 3 \- asset type id \- find by \-\> asset \= (select Property type from (“collateral” excel table) of the collateral id.on which we are working on). And then find corresponding asset type id of asset in CERSAI2.0\_Masters excel. In ASSET\_TYPE\_ID sheet**  
 **Column 4 \- by default (“Others”)**  
**Column 5 \-  get value from req details excel**  
 **Column 6 \- by default (NA)**  
 **Column 7 \- asset\_description (get value from req details table)**  
 **Column 8 \- Survey\_Number (get value from req details table)**  
**Column 9 \- Plot\_Number (get value from req details table)**  
 **Column 10 \- Asset\_Area (get value from req details table)**  
 **Column 11 \-  Asset\_Area\_Unit (get value from req details table)**  
 **Column 12 \-  House Number (get value from req details table)**  
**Column 13 \-  by default (blank)**  
 **Column 14 \-  optional \[if present in address in req details\]**  
**Column 15 \-  Project Name  (if present in req details table, fill it else leave it)**  
**Column 16- Street\_Name (get value from req details table)**  
**Column 17 \- pocket \- optional \[if present in address in req details\]**

 **for column 19 \- 23 \-\>**  
    
 **1\. address \= (select Full Address from (“collateral” excel table) of the collateral id. On which we are working on).**  
 **2\. Fetch PIN code from address**  
**3\. go to crs2\_addrss\_mstr\_pincd\_analysis\_20260416 excel \-\> effective\_master sheet**  
**4\. Search PIN code in this  excel sheet**  
**5\. Find corresponding District Code in this excel sheet**  
**6\. Find corresponding State Name to PIN code in this excel sheet**  
**7\. Go to CERSAI2.0\_Masters excel and go to STATE sheet \-\> Find corresponding State Code of the State Name**  
**8\. Fetch city from address**  
**9\. Go to CERSAI2.0\_Masters excel and go to CITY sheet \-\> Find corresponding City Code of the (city and District Code)**  
 **(if there are multiple rows of same city in CITY sheet \-\> make sure you select City Code from row which have our District Code(found in step 5\)**  
**10\. Find locality from address**  
**Column 18 \- value \= (locality)**  
 **Column 19 \-  value \= (City Code) // maintain in proper format(for eg \- don’t mention E)**  
 **Column 20 \- value  \= (District Code)**  
**Column 21-  value  \= (State Code)**  
 **Column 22 \- Value \= (PIN code)**  
**Column 23 \-  by default (IND)**  
**Row 4 to Row (4+ no. of borrowers \-1 )**

 **in these rows, we will have to mostly get value from CKYC excel table. We need to join CKYC excel table and “Working Collateral id” table on application number to find relevant information of borrowers CKYC table**  
**Column 1 \- File header \- value(“BOR”)**  
 **Column 2 \- Serial\_Number (ith value from no. of borrowers)**  
 **Column 3 \- Borrower\_Type (get value by 1\. Finding type of borrower from ckyc table, then 2\. find corresponding Borrower/Asset Owner Code from BORROWER\_ASSET\_OWNER\_TYPE sheet in CERSAI2.0\_Masters excel )**  
 **Column 4 \- Is\_Asset\_Owner (value from CKYC table) (1 if asset owner, else 0\)**  
**Column 5 \- by default (blank)**  
 **Column 6 \- by default (blank)**  
 **Column 7 \- Borrower\_Pan (get value by finding Pan number from CKYC table) when (application type \= non\_individual)**  
 **Column 8 \- by default (blank)**  
**Column 9 \- Borrower\_name (get value by finding Pan number from CKYC table) when (application type \= non\_individual)**  
 **Column 10 \- incorporation date  (get value by incorporation  date from CKYC table)**  
 **Column 11 \-  ??**  
 **Column 12 \- Individual Pan (get value by finding Pan number from CKYC table) when (application type \=individual)**  
**Column 13 \-  by default (blank)**  
 **Column 14 \- ??**  
**Column 15 \-   customer name(get value by finding customer name from CKYC table) when (application type \=individual)**  
**Column 16- ??**  
**Column 17 \- Date of Birth  Individual Pan (get value by finding Date of Birth from CKYC table) when (application type \=individual)**  
 **Column 18 \- Mobile No  Individual Pan (get value by finding Mobile No from CKYC table) when (application type \=individual)**  
**Similar to how we entered address in Row 3 for immovable asset, we have to add address of each borrower in their respective rows \-\>**

 **Column 20 \- value \= address of respective borrower (first 100 alphanumeric characters)**   
**Column 21 \-  value \=  if address of respective borrower is greater than 100 character, write remaining address here. Otherwise leave it blank. Max length \= 100**  
**Column 22 \- value \= if address of respective borrower is greater than 200 characters and doesn’t fit in even column 21, write remaining address here. Otherwise leave it blank. Max length of column \= 100**  
 **Column 23 \-  value \= (City Code) (leave blank for now. We will build its logic later)**  
 **Column 24 \- value  \= (District Code)**  
**Column 25-  value  \= (State Code)**  
 **Column 26 \- Value \= (PIN code)**  
**Column 27 \-  by default (“IN”)**  
**Values from Legal document \-**

 **si\_type, cts no, property owner, area,**   
**For Last Column \-**

 **Column 1 \= File header \- value(“CHG”)**  
**Column 2 \= by default(blank)**  
**Column 3 \= by default (“ANAND RATHI GLOBAL FINANCE LIMITED”)**  
**Column 4= by default(Express Zone B Wing 5th Floor Western Express Highway Opposite Oberoi Mall Malad East Mumbai 400063”)**  
**Column 5-8 \= by default (blank)**  
**Column 9 \= by default(“62995”)**  
**Column 10 \= by default(“10680”)**  
**Column 11 \= by default(“21”)**  
**Column 12 \= by default(“400063”)**  
**Column 13 \= by default(“IND”)**  
Query-\> is PAN compulsory, gender, father name  
If two flats are mentioned in full address but have single collateral id, then how will we resolve it.  
