Read_CUE_from_HM_Box<-function(CUE_filepath=NA)
{
#Load in library to read in excel docs (.xlsx)
#install.packages(xlsx)
library(xlsx)

#####################################################################################################
###Load the CUEs/IPAR files##########################################################################
#####################################################################################################

if(is.na(CUE_filepath)){
#create a filepath to where all the excel doc/CUE sheets are stored
file_path<-"C:\\Users\\Robert Chapman\\Box\\PCAR\\07. PCAR Measures\\*.*"

#All the user to choose the CUE files (in case there are files in the directory which are note CUEs)
CUE_list<-choose.files(default=file_path)
}
else
{CUE_list<-CUE_filepath} 
###Declare static text variables to be used in the loop
item_ID_list<-c("Toolbox.Variable.Name", "New.ID.","Variable.Name","VariableName")
a_var<-"Calibration.Statistics.Discrimination..slope."                                                                
threshold_parameters<-paste0("Calibration.Statistics.Location...Threshold.",1:6)
NCAT=integer()

#for loop to cycle through the individual CUE files
for(i in 1:length(CUE_list))
{
  j=1; #counter for which row we're looking at
  Raw_CUE_xlsx<-read.xlsx(CUE_list[i],j) #attempt to read in the excel formatted CUE sheet at Row J
  #if we don't find the slope variable in the list of variables read in from the CUE sheet, then cut one Row and re-read
  while(!any(names(Raw_CUE_xlsx)==a_var))
  {
    j=j+1
    Raw_CUE_xlsx<-read.xlsx(CUE_list[i],1, startRow = j)
  }
  
  #compare list of names available in the excel CUE sheet against the known list of "variable names", and return the matching name
  item_ID<- c(item_ID_list,names(Raw_CUE_xlsx))[duplicated(c(item_ID_list,names(Raw_CUE_xlsx)))]
  
  #return the number of items found in the CUE excel sheet (not counting those who don't that don't have a variable name)
  number_of_items<-length(which(!is.na(Raw_CUE_xlsx[a_var])==TRUE))
  
  #figure out the number of parameters
  for(k in 1:length(threshold_parameters)){
    if(!any(names(Raw_CUE_xlsx)==threshold_parameters[k]))
    {NCAT=k;break;}
    if(anyNA(Raw_CUE_xlsx[1:number_of_items,threshold_parameters[k]]))
    {NCAT=k;break;}
  }
  
  ipar<-Raw_CUE_xlsx[1:number_of_items,c("Item.Stem","Responses",a_var,threshold_parameters[1:(NCAT-1)])]
  rownames(ipar)<-Raw_CUE_xlsx[1:number_of_items,item_ID]  
  names(ipar)<-c("Stem", "Responses", "a", paste0("cb",seq(from=1,to=(NCAT-1))))  
  ipar$NCAT<-c(rep(NCAT, times=number_of_items))
}
return(ipar)
}
