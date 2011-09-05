%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
% Do not modify this file since it was automatically generated from:
% 
%  rtf.R
% 
% by the Rdoc compiler part of the R.oo package.
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

 \name{addTable.RTF}
\alias{addTable.RTF}
\alias{RTF.addTable}
\alias{addTable.RTF}
\alias{addTable,RTF-method}


 \title{Add a table to the RTF document}

 \description{
	Add a table to the RTF document.
 }

 \usage{\method{addTable}{RTF}(this, dat, col.widths=NULL, font.size=NULL, row.names=FALSE, NA.string="-", ...)}

 \arguments{
 	\item{this}{An RTF object.}
 	\item{dat}{A matrix, data frame, or table.}
	\item{col.widths}{A \code{\link[base]{vector}} of column widths in inches \bold{optional}.}
	\item{font.size}{Font size in points \bold{optional}.}
	\item{row.names}{Boolean argument to include row names in tables \bold{optional}.}
	\item{NA.string}{A character to replace NA values in the table.}
 	\item{...}{Not used.}
 }

 \value{
 	The new name (invisible).
 }

 \examples{
 rtf<-RTF("test.rtf.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 addTable(rtf,as.data.frame(head(iris)),font.size=10,row.names=FALSE,NA.string="-",
          col.widths=rep(1,5))
 done(rtf)
 }

 \author{Michael E. Schaffer}

 \seealso{
 	For more information see \code{\link{RTF}}.
 }
\keyword{internal}
\keyword{methods}
