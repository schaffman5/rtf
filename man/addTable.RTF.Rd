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


 \title{Insert a table into the RTF document}

 \description{
	Insert a table into the RTF document.
 }

 \usage{\method{addTable}{RTF}(this, dat, col.widths=NULL, col.justify=NULL, header.col.justify=NULL, font.size=NULL, row.names=FALSE, NA.string="-", space.before=NULL, space.after=NULL, ...)}

 \arguments{
 	\item{this}{An RTF object.}
 	\item{dat}{A matrix, data frame, or table.}
	\item{col.widths}{A \code{\link[base]{vector}} of column widths in inches. \bold{optional}.}
	\item{col.justify}{A single value or \code{\link[base]{vector}} of column justifications ('L', 'R', 'C', or 'J' for Left, Right, Center, and Justify, respectively). \bold{optional}.}
	\item{header.col.justify}{A single value or \code{\link[base]{vector}} of table header column justifications ('L', 'R', 'C', or 'J' for Left, Right, Center, and Justify, respectively). \bold{optional}.}
	\item{font.size}{Font size in points. \bold{optional}.}
	\item{row.names}{Boolean argument to include row names in tables. \bold{optional}.}
	\item{NA.string}{A character string to replace NA values in the table.}
	\item{space.before}{Space before each row (in inches). \bold{optional}.}
	\item{space.after}{Space after each row (in inches). \bold{optional}.}
 	\item{...}{Not used.}
 }

 \examples{
 rtf<-RTF("test_addTable.doc",width=8.5,height=11,font.size=10,omi=c(1,1,1,1))
 addTable(rtf,as.data.frame(head(iris)),font.size=10,row.names=FALSE,NA.string="-",
          col.widths=rep(1,5))

 tab<-table(iris$Species,floor(iris$Sepal.Length))
 names(dimnames(tab))<-c("Species","Sepal Length")
 addTable(rtf,tab,font.size=10,row.names=TRUE,NA.string="-",col.widths=c(1,rep(0.5,4)) )

 done(rtf)
 }

 \author{Michael E. Schaffer}

 \seealso{
 	For more information see \code{\link{RTF}}.
 }


\keyword{methods}
