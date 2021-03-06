#' Open an existing Echoview file (.EV) via COM scripting.
#' 
#' This function opens an existing Echoview (.EV) file using COM scripting.  
#' @param EVAppObj An EV application COM object arising from the call COMCreate('EchoviewCom.EvApplication')
#' @param fileName An Echoview file path and name.
#' @return a list object with two elements.  $EVFile: EVFile COM object, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVOpenFile(EVAppObj,'~\\example1.EV')
#'}
EVOpenFile=function(EVAppObj,fileName){
  nbrFilesOpen=EVAppObj$EvFiles()$Count()
  msg=paste(Sys.time(),' : There are currently ',nbrFilesOpen, 
            ' EV files open in the EV application.',sep='')
  message(msg[1])
  chkAlreadyOpen=EVAppObj[['EvFiles']]$FindByFileName(fileName)
  if(!is.null(chkAlreadyOpen))
  {
    msg=c(msg,paste(Sys.time(),' : File already open. EVOpenFile() 
                    returning existing EV file in EVFile object ',
                    fileName,sep=''))
    message(msg[2])
    return(list(EVFile=chkAlreadyOpen,msg=msg))
  }
    
  msg=c(msg,paste(Sys.time(),' : Opening ',fileName,sep=''))
  message(msg[2])
  EVFile=EVAppObj$OpenFile(fileName)
  dF=EVAppObj$EvFiles()$Count()-nbrFilesOpen #nbr of files now open.
  if(dF!=0 | dF!=1) msgTMP='Failed to open EV file, unknown error'
  if(dF==1) msgTMP='Opened EV file: '
  if(dF==0) msgTMP='Check filename. Failed to open EV file: '
  msg=c(msg,paste(Sys.time(),' : ',msgTMP,' ',fileName,sep=''))
  message(msg[3])
  invisible(list(EVFile=EVFile,msg=msg))
}


#' Saves an open Echoview file (.EV) via COM scripting.
#' 
#' This function saves an existing Echoview (.EV) file using COM scripting.  
#' @param EVFile An Echoview file COM object
#' @return a list object with two elements.  $chk: Boolean check indicating if the file was successfully saved; $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#'@examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVOpenFile(EVAppObj,'~\\example1.EV')$EVFile
#'EVSaveFile(EVFile)
#'}
EVSaveFile=function(EVFile){
  chk=EVFile$Save()
  msg=paste(Sys.time(),' : ',ifelse(chk,'Saved','Failed to save'),' ',EVFile$FileName(),sep='')
  if(chk) message(msg) else warning(msg)
  invisible(list(chk=chk,msg=msg))}


#' Performs save as operation on an open Echoview file (.EV) via COM scripting.
#' 
#' This function performs a save as operation on an existing Echoview (.EV) file using COM scripting.  
#' @param EVFile An Echoview file COM object
#' @param fileName An Echoview file path and name.
#' @return a list object with two elements.  $chk: Boolean check indicating if the file was successfully saved; $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVSaveFile}} \code{\link{EVCloseFile}} 
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVOpenFile(EVAppObj,'~\\example1.EV')$EVFile
#'EVSaveAsFile(EVFile=EVFile,fileName='~\\example1R1.EV')
#'}
EVSaveAsFile=function(EVFile,fileName){
  chk=EVFile$SaveAs(fileName)
  msg=paste(Sys.time(),' : ',ifelse(chk,'Saved','Failed to save'),' ',EVFile$FileName(),sep='')
  if(chk) message(msg) else warning(msg)
  invisible(list(chk=chk,msg=msg))}

#' Closes an open Echoview file (.EV) via COM scripting.
#' 
#' @param EVFile An Echoview file COM object
#' @return a list object with two elements.  $chk: Boolean check indicating if the file was successfully closed; $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm}
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVOpenFile(EVAppObj,'~\\example1.EV')$EVFile
#'EVCloseFile(EVFile)
#'}
EVCloseFile=function(EVFile){
  fn=EVFile$FileName()
  chk=EVFile$Close()
  msg=paste(Sys.time(),' : ',ifelse(chk,'Closed','Failed to close'),' ',fn,sep='')
  if(chk) message(msg) else warning(msg)
  invisible(list(chk=chk,msg=msg))}

#' Creates a new Echoview file (.EV) via COM scripting.
#' 
#' Creates a new Echoview file (.EV) via COM scripting which may be created from a template file if available
#' @param EVAppObj An EV application COM object arising from the call COMCreate('EchoviewCom.EvApplication')
#' @param templateFn full path and filename for an Echoview template
#' @return a list object with two elements.  $EVFile: EVFile COM object for the newly created Echovie file, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVNewFile(EVAppObj)$EVFile
#'}
EVNewFile=function(EVAppObj,templateFn=NULL){
  #ARGS: suggest users specify full path for any template file.
  if(is.null(templateFn)){
    EvFile=EVAppObj$NewFile()
    msgV=paste(Sys.time(),' : ','New blank Echoview file created',sep='')
  }
  if(is.character(templateFn)){
    msgV=paste(Sys.time(),' : Attempting to create an Echoview file from template ',templateFn,sep='')
    message(msgV)
    EvFile=EVAppObj$NewFile(templateFn)
    nbrVars=EvFile[['Variables']]$Count()
    if(nbrVars==0){msg=paste(Sys.time(),' : Either an incorrect template filename or template contains no variables',sep='')
                   warning(msg)
                   msgV=c(msgV,msg)} else {
                     msg= paste(Sys.time(),' : Echoview file created from template ',templateFn,sep='')
                     message(msg)
                     msgV=c(msgV,msg)
                   }
    
  } 
  if(!is.character(templateFn) & !is.null(templateFn)) {
    msgV=paste(Sys.time(),' : Incorrect ARG templateFn specification in EVNewFile()',sep='')
    warning(msgV)
    return(msgV)}
  invisible(list(EVFile=EvFile,msg=msgV))}

#' Creates a new Echoview fileset via COM scripting
#' 
#' @param EVFile An Echoview file COM object
#' @param filesetName Echoview fileset name to create
#' @return a list object with two elements.  $fileset: created fileset COM object, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVCreateNew}}
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVOpenFile(EVAppObj,'~\\example1.EV')$EVFile
#'EVCreateFileset(EVFile=EVFile,filesetName='example')
#'}
EVCreateFileset=function(EVFile,filesetName){
  msgV=paste(Sys.time(),' : Creating fileset called ',filesetName,' in ',EVFile$FileName(), sep='')
  message(msgV)
  allFilesets=EVFile[["Filesets"]] 
  chk=allFilesets$Add(filesetName)
  if(chk){
    msg=paste(Sys.time(),' : Successfully created fileset called ',filesetName,' in ',EVFile$FileName(), sep='')
    message(msg)
    msgV=c(msgV,msg)
    invisible()
  } else {}
}

#' Finds an Echoview fileset in an Echoview file via COM scripting
#' 
#' @param EVFile An Echoview file COM object
#' @param filesetName Echoview fileset name to find
#' @return a list object with two elements.  $fileset: found fileset COM object, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVCreateNew}} \code{\link{EVCreateFileset}}
#' @examples
#'\dontrun{
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVOpenFile(EVAppObj,'~\\example1.EV')$EVFile
#'EVFindFilesetByName(EVFile=EVFile,filesetName='example')
#'}
EVFindFilesetByName=function(EVFile,filesetName) {
  msgV=paste(Sys.time(),' : Searching for fileset name ',filesetName,' in ',EVFile$FileName(), sep='')
  message(msgV)
  allFilesets=EVFile[["Filesets"]]
  filesetObj=allFilesets$FindByName(filesetName)

  if(class(filesetObj)[1]=="COMIDispatch"){
    msg=paste(Sys.time(),' : Found the ',filesetObj$Name(),' fileset in ',EVFile$FileName(), sep='')
    message(msg)
    msgV=c(msgV,msg)
    return(list(filesetObj=filesetObj,msg=msgV))} else{
    msg=paste(Sys.time(),' : The fileset called ',filesetName, ' not found in ',EVFile$FileName(), sep='')
    warning(msg)
    msgV=c(msgV,msg)
    return(list(filesetObj=NULL,msg=msgV))
    }  
}


#' Adds raw data files to an open Echoview file (.EV) via COM scripting.
#' 
#' Adds raw data files to an open Echoview file (.EV) via COM scripting.  The function assumes the Echoview fileset name already exists.
#' @param EVFile An Echoview file COM object
#' @param filesetName Echoview fileset name
#' @param dataFiles vector of full path and name for each data file.
#' @return a list object with two elements.  $nbrFilesInFileset: number of raw data files in the fileset, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVCreateNew}}
#' @examples
#'\dontrun{
#'filenamesV=c('~\\exampleData\\ek60-1.raw','~\\exampleData\\ek60-2.raw','~\\exampleData\\ek60-3.raw')
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVNewFile(EVAppObj,templateFn="~\\Example-template")$EVFile
#'EVAddRawData(EVFile=EVFile,filesetName='EK60',dataFiles=filenamesV)
#'}
EVAddRawData=function(EVFile,filesetName,dataFiles)
{
  #get number of raw data files currently in "fileset.name" fileset
  nbr.of.raw.in.fileset.pre=destination.fileset[["DataFiles"]]$Count()
  #add new files
  msgV=paste(Sys.time(),' : Adding data files to EV file ',sep='')
  message(msgV)
  for(i in 1:length(dataFiles)){
    destination.fileset[["DataFiles"]]$Add(dataFiles[i]) 
    msg=paste(Sys.time(),' : Adding ', dataFiles[i],' to fileset name ',filesetName,sep='')
    message(msg)
    msgV=c(msgV,msg)
  }
  
  nbr.of.raw.in.fileset=destination.fileset[["DataFiles"]]$Count()
  if((nbr.of.raw.in.fileset-nbr.of.raw.in.fileset.pre)!=length(dataFiles)){
    msg=paste(Sys.time(),' : Number of candidate to number of added file mismatch',sep='')
    msgV=c(msgV,msg)
    warning(msg)
  }
  
  invisible(list(nbrFilesInFileset=nbr.of.raw.in.fileset,msg=msgV))
} #EVAddRawData

#' Creates a new Echoview (.EV) file and adds raw data files to it via COM scripting.
#' 
#' Creates a new Echoview (.EV) file and adds raw data files to it via COM scripting.  Works well when populating an existing Echoview template file with raw data files.  The newly created Echoview file will remain open in Echoview and can be accessed via the $EVFile objected returned by a successful call of this function.
#' @param EVAppObj An EV application COM object arising from the call COMCreate('EchoviewCom.EvApplication')
#' @param EVFileName Full path and filename of Echoview (.EV) file to be created.
#' @param filesetName Echoview fileset name
#' @param dataFiles vector of full path and name for each data file.
#' @return a list object with two elements.  $EVFile: EVFile COM object for the newly created Echoview file, and $msg: message for processing log. 
#' @keywords Echoview COM scripting
#' @export 
#' @references \url{http://support.echoview.com/WebHelp/Echoview.htm/}
#' @seealso \code{\link{EVNewFile}}  \code{\link{EVAddRawData}}  \code{\link{EVCloseFile}}
#' @examples
#'\dontrun{
#'filenamesV=c('~\\exampleData\\ek60-1.raw','~\\exampleData\\ek60-2.raw','~\\exampleData\\ek60-3.raw')
#'EVAppObj=COMCreate('EchoviewCom.EvApplication')
#'EVFile=EVCreateNew(EVAppObj,templateFn="~\\Example-template",filesetName='EK60',dataFiles=filenamesV)
#'}
EVCreateNew=function(EVAppObj,templateFn=NULL,EVFileName,filesetName,
                     dataFiles)
{
  #20100726 creates ev file from a template and populates the ev file with data.
  #NB this function will save and close the EV file once raw data files are added.
  #REQUIRES RDCOMClient package; EVAddRawData
  #INPUTS: EvApp = EV program object comCreateObject("EchoviewCom.EvApplication")
  #        template.fn = template filename located in the templates EV directory.
  #        ev.op.dir = directory to save EV file to
  #        ev.op.fn = EV filename.
  #        fileset.anme = fileset name in template.fn.
  #         dataFiles = vector of character strings for raw data location e.g.
  #                     "D:/JC037/SW/T01/JC037-D20090822-T102525.raw" 
  msgV=paste(Sys.time(),' : Creating new EV file',sep='')
  message(msgV)
  
  EVFile=EVNewFile(EVAppObj=EVAppObj,templateFn=templateFn)
  msgV=c(msgV,EVFile$msg)
  EVFile=EVFile$EVFile
  msgV=c(msgV,EVAddRawData(EVFile=EVFile,filesetName=filesetName,dataFiles=dataFiles)$msg)
  msgV=c(msgV,EVSaveAsFile(EVFile=EVFile,EVFn=paste(EVOpDir,EVOpFn,sep=""))$msg)
  #msgV=c(msgV,EVCloseFile(EVFile=EVFile)$msg)
  
  return(list(EVFile=EVFile,msg=msgV))
} #end -EVcreateNew



EVminThresholdSet<-function(varObj,thres){
  varDat=varObj[["Properties"]][["Data"]]
  preThresApplyFlag<-varDat$ApplyMinimumThreshold()
  varDat[['ApplyMinimumThreshold']]<-TRUE
  postThresApplyFlag<-varDat$ApplyMinimumThreshold()
  if(postThresApplyFlag){msg<-paste(Sys.time(),' : Apply minimum threshold flag set to TRUE in ',
                                    varObj$Name(),sep='')
                         message(msg)
  }else{
    msg<-paste(Sys.time(),' : Failed to set minimum threshold flag in ',
               varObj$Name(),sep='')
    stop(msg)}
  #set threshold value
  preMinThresVal=varDat$MinimumThreshold()
  varDat[['MinimumThreshold']]<-thres
  postMinThresVal=varDat$MinimumThreshold()
  if(postMinThresVal==thres){
    msg2<-paste(Sys.time(),' : Minimum threshold successfully set to ',thres,' in ',
                varObj$Name(),sep='')
    message(msg2)
    msgV=c(msg,msg2)
  }else{
    msg2<-paste(Sys.time(),' : Failed to set minimum threshold in ',
                varObj$Name(),sep='')
    stop(msg2)
    msgV=c(msg,msg2)
  }
  #now try with display threshold
  varDisp=varObj[["Properties"]][['Display']]
  preDisplayThres<-varDisp$ColorMinimum()
  varDisp[['ColorMinimum']]<-thres
  postDisplayThres<-varDisp$ColorMinimum()
  if(thres==postDisplayThres){
    msg=paste(Sys.time(),' : Display threshold also changed to',thres,' dB re 1m^-1',sep='')
  message(msg)}  else{
    msg<-paste(Sys.time(),' : Failed to change display threshold to ',thres,' dB re 1m^-1 \n',sep='')
    warning(msg)}
  msgV=c(msg,msgV)
  
  return(list(thresholdSettings=c(preThresApplyFlag=preThresApplyFlag, 
                                  preMinThresVal=preMinThresVal,
                                  postThresApplyFlag=postThresApplyFlag,
                                  preMinThresVal=preMinThresVal,postMinThresVal=postMinThresVal),msg=msgV))
}  
#EVminThresholdSet(varObj=varObj,thres=(-60))

EVSchoolsDetSet=function(EVFile,varObj,distanceMode,
                         maximumHorizontalLink,
                         maximumVerticalLink,
                         minimumCandidateHeight,
                         minimumCandidateLength,
                         minimumSchoolHeight,
                         minimumSchoolLength){
  # 20120309 set schools detection parameters.
  #ARGS: EvFile = EV file object;
  # var.nbr = EV virtual variable object number to change the threshold of
  #pars = parameter vector for the schools detection module. vector is structured as follows:
  #     pars[1] = Distance Mode;
  # pars[2] = Maximum horizontal link distance (m);
  #   pars[3] =  Maximum vertical link (m);
  #    pars[4] = MinimumCandidateHeight (m);
  #     pars[5] = MinimumCandidateLength (m);
  #     pars[6] = MinimumSchoolHeight (m);
  #     pars[7] = MinimumSchoolLength (m)
  #returns dataframe of current parameters, revised parameters
  setVec=c(maximumHorizontalLink, maximumVerticalLink, minimumCandidateHeight, 
           minimumCandidateLength, minimumSchoolHeight,minimumSchoolLength)
  if(!all(is.numeric(setVec)))
    stop('Non-numeric ARG in school detection distance settings')
  #get schools object from the current EvFile properties:  
  school.obj=EVFile[["Properties"]][["SchoolsDetection2D"]] 
  #get current school detection parameters
  pre.distmode=school.obj[['DistanceMode']]
  pre.maxhzlink=school.obj[["MaximumHorizontalLink"]]
  pre.maxvtlink=school.obj[["MaximumVerticalLink"]]
  pre.mincandHt=school.obj[["MinimumCandidateHeight"]]
  pre.mincandLen=school.obj[["MinimumCandidateLength"]]
  pre.minSchoolHt=school.obj[["MinimumSchoolHeight"]]
  pre.minSchoolLen=school.obj[["MinimumSchoolLength"]]
  preSettingDistances=c(pre.maxhzlink=pre.maxhzlink,
                        pre.maxvtlink=pre.maxvtlink,pre.mincandHt=pre.mincandHt,pre.mincandLen=pre.mincandLen,
                        pre.minSchoolHt=pre.minSchoolHt,pre.minSchoolLen=pre.minSchoolLen)
  #set school parameters
  school.obj[["DistanceMode"]]<-distanceMode
  school.obj[["MaximumHorizontalLink"]]<-maximumHorizontalLink
  school.obj[["MaximumVerticalLink"]]<- maximumVerticalLink
  school.obj[["MinimumCandidateHeight"]]<-minimumCandidateHeight
  school.obj[["MinimumCandidateLength"]]<-minimumCandidateLength
  school.obj[["MinimumSchoolHeight"]]<-minimumSchoolHeight
  school.obj[["MinimumSchoolLength"]]<- minimumSchoolLength
  #check settings have been applied
  #get current (post change) school detection parameters
  post.distmode=school.obj[["DistanceMode"]]
  post.maxhzlink=school.obj[["MaximumHorizontalLink"]]
  post.maxvtlink=school.obj[["MaximumVerticalLink"]]
  post.mincandHt=school.obj[["MinimumCandidateHeight"]]
  post.mincandLen=school.obj[["MinimumCandidateLength"]]
  post.minSchoolHt=school.obj[["MinimumSchoolHeight"]]
  post.minSchoolLen=school.obj[["MinimumSchoolLength"]]
  postSettingDistances=c(post.maxhzlink=post.maxhzlink,
                         post.maxvtlink=post.maxvtlink,post.mincandHt=post.mincandHt,post.mincandLen=post.mincandLen,
                         post.minSchoolHt=post.minSchoolHt,post.minSchoolLen=post.minSchoolLen)
  
  if(post.distmode!=distanceMode){
    msg<-paste(Sys.time()," : Failed to set distance mode in schools detection",sep="")
    invisible(msg)
    stop(msg)
  }
  setDiff=which(postSettingDistances != setVec)
  if(length(setDiff)>0){
    msg<-paste(Sys.time(),' : Failed to set schools detection parameters: ',
               paste(names(PostSettingDistances)[setDiff],collapse=', '),sep='')
    invisible(msg)
    stop(msg)
  }else{
    msg<-paste(Sys.time(),' : Set schools detection parameters: Distance mode = ',post.distmode,' ',
               paste(names(postSettingDistances),'=',postSettingDistances,collapse='; '),sep='')  
    message(msg)
  }
  out=list(pre.distmode=pre.distmode,preSettingDistances=preSettingDistances,
           post.distmode=post.distmode,postSettingDistances=postSettingDistances,msg=msg)
  return(out)
}
#EVSchoolsDetSet(EVFile,varObj,distanceMode="GPS distance",
#                maximumHorizontalLink=10,#m
#                maximumVerticalLink=5,#m
#                minimumCandidateHeight=1,#m
#                minimumCandidateLength=10,#m
#                minimumSchoolHeight=2,#m
#                minimumSchoolLength=15)


EVAcoVarNameFinder=function(EVFile,acoVarName){
  
  obj=EVFile[["Variables"]]$FindByName(acoVarName)
  if(is.null(obj))
    obj=EVFile[["Variables"]]$FindByShortName(acoVarName)
  if(is.null(obj)){
    msg=paste(Sys.time(),' : Variable not found ',acoVarName,sep='')
    warning(msg)
    obj=NULL}else{
      msg=paste(Sys.time(),' : Variable found ',acoVarName,sep='') 
      message(msg)
    }
  return(list(EVVar=obj,msg=msg))  
}
#EVAcoVarNameFinder(EVFile,acoVarName='120H Sv mri 0-250m 7x7 convolution')

EVRegionClassFinder=function(EVFile,regionClassName){
  obj=EVFile[["RegionClasses"]]$FindByName(regionClassName)
  if(is.null(obj)){
    msg=paste(Sys.time(),' : Region class not found -',regionClassName,sep='')
    warning(msg)
    obj=NULL}else{
      msg=paste(Sys.time(),' : Region class found -',regionClassName,sep='') 
      message(msg)
    }
  return(list(regionClass=obj,msg=msg))
}

#EVRegionClassFinder(EVFile,regionClassName='krill_swarm')

EVDeleteRegionClass=function(EVFile,regionClassCOMObj){
  classChk=class(regionClassCOMObj)[1]
  if(classChk!='COMIDispatch'){
    msg=paste(Sys.time(),' : attempted to pass non-COM object in ARG regionClassCOMObj',sep='')
    stop(msg)
  }
  nbrRegionsPre=EVFile[['Regions']]$Count()
  del=EVFile[['Regions']]$DeleteByClass(regionClassCOMObj)
  nbrRegionsDel=nbrRegionsPre-EVFile[['Regions']]$Count()
  if(is.null(del)){
    msg=paste(Sys.time(),' : Regions in region class',regionClassCOMObj$Name(),' not deleted',sep='')
    warning(msg)}else{
      msg=paste(Sys.time(),' : Regions in region class, ',regionClassCOMObj$Name(),', deleted. ',nbrRegionsDel ,' individual regions deleted.',sep='') 
      message(msg)
    }
  invisible(msg)
}
#EVDeleteRegionClass(EVFile=EVFile,regionClassCOMObj=regObj)


EVSchoolsDetect=function(
  EVFile,
  acoVarName,
  outputRegionClassName,
  deleteExistingRegions,
  distanceMode,
  maximumHorizontalLink,
  maximumVerticalLink,
  minimumCandidateHeight,
  minimumCandidateLength,
  minimumSchoolHeight,
  minimumSchoolLength,
  dataThreshold){
  
  #find acoustic variable:
  varObj<-EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=varObj$msg
  varObj=varObj$EVVar
  if(is.null(varObj)){
    msgV=c(msgV,paste(Sys.time(),' : Stopping schools detection, acoustic variable not found',sep=''))
    message(msgV)
    return(list(nbDetections=NULL,msg=msgV))}
  #find region class
  regObj=EVRegionClassFinder(EVFile=EVFile,regionClassName=outputRegionClassName)
  msgV=c(msgV,regObj$msg)
  regObj=regObj$regionClass
  if(is.null(regObj)){
    msgV=c(msgV,paste(Sys.time(),' : Stopping schools detection, region class not found',sep=''))
    message(msgV[2])
    return(list(nbDetections=NULL,msg=msgV))}
  #handling exisiting regions:
  if(deleteExistingRegions){
    msgV=c(msgV,EVDeleteRegionClass(EVFile=EVFile,regionClassCOMObj=regObj))} else {
      msg=paste(Sys.time(),' : adding detected regions those existing in region class ', regObj$Name(),sep='')
      message(msg)
      msgV=c(msgV,msg)}
  #set threshold
  thresRes=EVminThresholdSet(varObj=varObj,thres= dataThreshold)
  msgV=c(msgV,thresRes$msg)
  #set schools detection parameters
  schoolDetSet<-EVSchoolsDetSet(EVFile=EVFile,varObj=varObj,distanceMode=distanceMode,
                                maximumHorizontalLink=maximumHorizontalLink,
                                maximumVerticalLink=maximumVerticalLink,
                                minimumCandidateHeight=minimumCandidateHeight,
                                minimumCandidateLength=minimumCandidateLength,
                                minimumSchoolHeight=minimumSchoolHeight,
                                minimumSchoolLength=minimumSchoolLength)
  msgV=c(msgV,schoolDetSet$msg)
  msg=paste(Sys.time(),' : Detecting schools in variable ',varObj$name(),sep='')
  message(msg)
  msgV=c(msgV,msg)
  nbrDetSchools=varObj$DetectSchools(regObj$Name())
  if(nbrDetSchools==-1){
    msg=paste(Sys.time(),' : Schools detection failed.')
    warning(msg)
  } else {
    msg=paste(Sys.time(),' : ',nbrDetSchools ,' schools detected in variable ',varObj$Name(),sep='')
    message(msg)}
  msgV=c(msgV,msg)
  out=list(nbrOfDetectedschools=nbrDetSchools,thresholdData=thresRes,schoolsSettingsData=schoolDetSet,msg=msgV)
  return(out)
}


# EVSchoolsDetect(EVFile=EVFile,
#                 acoVarName='120H Sv mri 0-250m 7x7 convolution',
#                 outputRegionClassName='krill_swarm',
#                 deleteExistingRegions=TRUE,
#                 distanceMode="GPS distance",
#                 maximumHorizontalLink=15,#m
#                 maximumVerticalLink=5,#m
#                 minimumCandidateHeight=1,#m
#                 minimumCandidateLength=10,#m
#                 minimumSchoolHeight=2,#m
#                 minimumSchoolLength=15, #m
#                 dataThreshold=-80)


EVIntegrationByRegionsExport<- function(EVFile,acoVarName,regionClassName,exportFn,
                                        dataThreshold=NULL){
  
  acoVarObj<-EVAcoVarNameFinder(EVFile=EVFile,acoVarName=acoVarName)
  msgV=acoVarObj$msg
  acoVarObj=acoVarObj$EVVar
  EVRC=EVRegionClassFinder(EVFile=EVFile,regionClassName=regionClassName)
  msgV=c(msgV,EVRC$msg)
  RC=EVRC$regionClass
  if(is.null(dataThreshold)){
    msg=paste(Sys.time(),' : Removing minimum data threshold from ', acoVarName, sep='')
    message(msg)
    msgV=c(msgV,msg)
    varDat=acoVarObj[["Properties"]][["Data"]]
    varDat[['ApplyMinimumThreshold']]<-FALSE
  } else {
    msg=EVminThresholdSet(varObj=acoVarObj,thres= dataThreshold)$msg
    message(msg)
    msgV=c(msgV,msg)}
  
  msg=paste(Sys.time(),' : Starting integration and export of ',regionClassName,sep='')
  message(msg)
  success<-acoVarObj$ExportIntegrationByRegions(exportFn,RC)
  if(success){
    msg=paste(Sys.time(),' : Successful integration and export of ',regionClassName,sep='')
    message(msg)
    msgV=c(msgV,msg)
  }else{
    msg<-paste(Sys.time(),' : Failed to integrate and/or export ',regionClassName,sep='')
    warning(msg)
    msgV=c(msgV,msg)
  } 
  invisible(list(msg=msgV))
}

#' Convert a Microsoft DATE object to a human readable date and time.
#' 
#' Time stamps in Echoview, such as start and end times of, for example, invididual regions, use the Microsoft DATE format.  This function converts the Microsoft DATE object to a human readable date and time.  NB no time zone is returned.
#' @param dateObj a Microsoft date object
#' @return time stamp in the format yyyy-mm-dd hh:mm:ss
#' @seealso www.echoview.com
msDATEConversion=function(dateObj){
  eTimeStamp=as.numeric(dateObj)
  eDate=as.Date(eTimeStamp,origin = "1899-12-30")
  second(eDate)<-ddays(eTimeStamp-floor(eTimeStamp))
  return(eDate)
}