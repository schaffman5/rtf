%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  rtf.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{addPageBreak.RTF}
\alias{addPageBreak.RTF}
\alias{RTF.addPageBreak}
\alias{addPageBreak.RTF}
\alias{addPageBreak,RTF-method}


 \title{Insert a page break into the RTF document optionally specifying new page settings}

 \description{
	Insert a page break into the RTF document optionally specifying new page settings.
 }

 \usage{\method{addPageBreak}{RTF}(this, width=8.5, height=11, omi=c(1, 1, 1, 1), ...)}

 \arguments{
 	\item{this}{An RTF object.}
 	\item{width}{New page width in inches. \bold{optional}.}
 	\item{height}{New page height in inches. \bold{optional}.}
 	\item{omi}{A \code{\link[base]{vector}} of page margins (botton, left, top, right) \bold{optional}.}
 	\item{...}{Not used.}
 }

 \examples{
 rtf<-RTF("test_addPageBreak.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 addPageBreak(rtf,width=11,height=8.5,omi=c(0.5,0.5,0.5,0.5))
 done(rtf)
 }

 \author{Michael E. Schaffer}

 \seealso{
 	For more information see \code{\link{RTF}}.
 }


\keyword{methods}
