<#-----------------------------------------------------------------------------------------# 
Author: Uzair Ansari 
 
Function: This script will download the KB articles from Microsoft update catalog 
          Input list of KB articles need to be provided 
 
Date: 25th June 2018 
 
Version 1.1 : Added OS filter 
 
Version: 1.0 
 
#-----------------------------------------------------------------------------------------#> 
 
 
<#-----------------------------------------------------------------------------------------# 
NOTES: 
1) By default this script will download all the KB articles present on the update 
   catalog website. If you want to download only specific updates like 'x64' or 
   only 'Cumulative' updates then you need to remove the # preceding the type you want 
   to download. For example if you want to download only Cumulative updates for x64 bit 
   machines then remove the # preceding "#$Platform = 'x64'" and #$Type = 'Cumulative'" 
 
 
2) Kindly provide the path where you wish to download the patches in $parentpath variable. 
   Kindly provide the text file path where KB article details are saved in $kblist variable. 
   KB details should be saved one below the other in text file as: 
    
   KB123456 
   KB567890 
    
#-----------------------------------------------------------------------------------------#> 
