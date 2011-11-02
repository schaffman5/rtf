%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  rtf.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{addSessionInfo.RTF}
\alias{addSessionInfo.RTF}
\alias{RTF.addSessionInfo}
\alias{addSessionInfo.RTF}
\alias{addSessionInfo,RTF-method}


 \title{Insert session information into the RTF document}

 \description{
	Insert session information into the RTF document.
 }

 \usage{\method{addSessionInfo}{RTF}(this, locale=TRUE, ...)}

 \arguments{
 	\item{this}{An RTF object.}
	\item{locale}{Output the locale.}
 	\item{...}{Not used.}
 }

 \details{
 	Exports session information to the RTF document in a similar
 }

 \examples{
 rtf<-RTF("test_addSessionInfo.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 addSessionInfo(rtf)
 done(rtf)
 }

 \author{Michael E. Schaffer}

 \seealso{
   For more information see \code{\link{RTF}}., \code{\link{sessionInfo}}.
 }
\keyword{internal}
\keyword{methods}
